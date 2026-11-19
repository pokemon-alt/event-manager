"""Workspace-scoped OpenAI settings and text-only profile edit planning."""
import asyncio
import http.client
import json
import os
import re

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from fastapi import HTTPException, Request

from event_state import coordinator_table_exists, read_json, read_key

MODEL = 'gpt-4.1-mini'
SETTINGS_PATH = '/api/event-coordinator/ai/settings'
PLAN_PATH = '/api/event-coordinator/ai/plan'
FIELDS = set('nickname phone firstName lastName street street2 street3 city state zip country '
             'billingFirstName billingLastName billingPhone billingStreet billingStreet2 '
             'billingStreet3 billingCity billingState billingZip billingCountry group notes'.split())
OPS = ('set', 'prefix', 'suffix', 'replace', 'trim', 'uppercase', 'lowercase', 'titlecase',
       'extract_email', 'copy', 'prefix_letters')
INSTRUCTIONS = """Translate the user's spreadsheet edit into a short sequence of operations on the
requested editable columns. Only the provided column IDs are available. Unprefixed address and
name fields are shipping fields. Operations run in order,
independently for each profile row. Use set for a literal (empty clears a cell), prefix/suffix
for literal text, replace for a case-sensitive literal find and replacement, trim for surrounding
whitespace, uppercase/lowercase/titlecase for capitalization, extract_email to keep the first
email in a cell (keeps the original when none is found), copy to copy from another available
field ID in value, and prefix_letters to prepend 1–8 random uppercase letters followed by a space
(value is the digit). Use an empty string for each unused find or value. Use at most 12 operations.
Return a short user-facing preview summary in message. If the instruction is ambiguous, requires
a row filter, needs real cell values, or cannot be fully expressed with these operations, return
no operations and explain the limitation or ask a clarifying question. Never invent substitute edits."""


def key_record(db):
    if db.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='event_coordinator_ai'").fetchone():
        return db.execute('SELECT credential FROM event_coordinator_ai WHERE slot=1').fetchone()
    return None


def key_context(resources):
    return ('event-coordinator/openai/v1:' + resources.identity).encode()


def save_credential(resources, value):
    with resources.database.connect() as db:
        db.execute('BEGIN IMMEDIATE')
        if value is None:
            if key_record(db):
                db.execute('DELETE FROM event_coordinator_ai WHERE slot=1')
            return
        # A missing encryption key with existing encrypted data must be restored.
        can_create = not key_record(db) and not coordinator_table_exists(db)
        key = read_key(resources.database.home, create=can_create)
        nonce = os.urandom(12)
        encrypted = nonce + AESGCM(key).encrypt(nonce, value.encode(), key_context(resources))
        db.execute('CREATE TABLE IF NOT EXISTS event_coordinator_ai '
                   '(slot INTEGER PRIMARY KEY CHECK(slot=1), credential BLOB NOT NULL)')
        db.execute('INSERT OR REPLACE INTO event_coordinator_ai(slot,credential) VALUES(1,?)', (encrypted,))


def load_credential(resources):
    with resources.database.connect() as db:
        record = key_record(db)
        if record is None:
            raise HTTPException(409, 'Add an OpenAI API key in Settings first.')
        blob = record['credential']
        key = read_key(resources.database.home)
        try:
            return AESGCM(key).decrypt(blob[:12], blob[12:], key_context(resources)).decode()
        except (ValueError, InvalidTag, TypeError):
            raise HTTPException(503, 'The saved API key could not be opened. Replace it in Settings.') from None


def object_schema(properties):
    return {'type': 'object', 'properties': properties, 'required': list(properties), 'additionalProperties': False}


def plan_schema(fields):
    string = {'type': 'string'}
    operation = object_schema({'field': {'type': 'string', 'enum': fields}, 'op': {'type': 'string', 'enum': list(OPS)},
                               'value': string, 'find': string})
    return object_schema({'operations': {'type': 'array', 'items': operation}, 'message': string})


def valid_plan(plan, fields):
    if not isinstance(plan, dict) or set(plan) != {'operations', 'message'}:
        return False
    operations, message = plan['operations'], plan['message']
    if not isinstance(message, str) or not message.strip() or len(message) > 1000:
        return False
    if not isinstance(operations, list) or len(operations) > 12:
        return False
    for item in operations:
        if not isinstance(item, dict) or set(item) != {'op', 'field', 'value', 'find'}:
            return False
        if not all(isinstance(v, str) and len(v) <= 500 for v in item.values()):
            return False
        field, op, value, find = (item[k] for k in ('field', 'op', 'value', 'find'))
        if field not in fields or op not in OPS or bool(find) != (op == 'replace'):
            return False
        if op in ('prefix', 'suffix') and not value or op == 'copy' and value not in fields:
            return False
        if op == 'prefix_letters' and not re.fullmatch('[1-8]', value):
            return False
        if op in ('trim', 'uppercase', 'lowercase', 'titlecase', 'extract_email') and value:
            return False
    return True


def request_plan(api_key, instruction, fields):
    payload = {'model': MODEL, 'store': False, 'instructions': INSTRUCTIONS, 'input': instruction,
               'max_output_tokens': 2400,
               'text': {'format': {'type': 'json_schema', 'name': 'profile_edits', 'strict': True,
                                   'schema': plan_schema(fields)}}}
    conn = http.client.HTTPSConnection('api.openai.com', timeout=35)
    try:
        conn.request('POST', '/v1/responses', body=json.dumps(payload).encode(),
                     headers={'Content-Type': 'application/json', 'Authorization': 'Bearer ' + api_key})
        response = conn.getresponse()
        if response.status in (401, 403):
            raise HTTPException(422, 'OpenAI denied access. Check the key and its model permissions in Settings.')
        if response.status == 429:
            raise HTTPException(429, 'OpenAI reached a rate or usage limit. Check your API billing or retry later.')
        if response.status != 200:
            raise HTTPException(502, 'OpenAI could not prepare the edit. Try again.')
        body = response.read(128 * 1024 + 1)
    except TimeoutError:
        raise HTTPException(504, 'OpenAI timed out. Try the preview again.') from None
    except (OSError, http.client.HTTPException):
        raise HTTPException(502, 'Could not connect to OpenAI. Try again.') from None
    finally:
        conn.close()
    if len(body) <= 128 * 1024:
        try:
            result = json.loads(body)
            if result['status'] == 'completed':
                parts = [part for output in result['output'] if output.get('type') == 'message'
                         for part in output['content']]
                if any(part.get('type') == 'refusal' for part in parts):
                    raise HTTPException(422, 'OpenAI declined this edit. Try a different description.')
                plan = json.loads(''.join(part['text'] for part in parts if part.get('type') == 'output_text'))
                if valid_plan(plan, fields):
                    return plan
        except (KeyError, TypeError, ValueError, AttributeError, RecursionError):
            pass
    raise HTTPException(502, 'OpenAI returned an incomplete or unsupported edit. Try a more specific description.')


def install(app, pinned):
    @app.get(SETTINGS_PATH)
    def status(request: Request):
        resources = pinned(request)
        with resources.database.connect() as db:
            return {'configured': key_record(db) is not None}

    @app.put(SETTINGS_PATH)
    async def save(request: Request):
        resources = pinned(request, True)
        body = await read_json(request, limit=8192)
        api_key = body.get('apiKey')
        if set(body) != {'apiKey'} or not isinstance(api_key, str) or not re.fullmatch(r'[!-~]{1,1024}', api_key.strip()):
            raise HTTPException(422, 'Enter an API key without spaces or line breaks.')
        save_credential(resources, api_key.strip())
        return {'configured': True}

    @app.delete(SETTINGS_PATH)
    def remove(request: Request):
        save_credential(pinned(request, True), None)
        return {'configured': False}

    @app.post(PLAN_PATH)
    async def plan(request: Request):
        resources = pinned(request, True)
        body = await read_json(request, limit=32768)
        fields, instruction = body.get('fields'), body.get('instruction')
        if set(body) != {'instruction', 'fields'} or not isinstance(instruction, str) or not 1 <= len(instruction.strip()) <= 2000:
            raise HTTPException(422, 'Describe the edit in 1 to 2,000 characters.')
        if not isinstance(fields, list) or not 1 <= len(fields) <= len(FIELDS) or not all(
                isinstance(field, str) and field in FIELDS for field in fields) or len(set(fields)) != len(fields):
            raise HTTPException(422, 'Choose editable text columns for this preview.')
        api_key = load_credential(resources)
        return await asyncio.to_thread(request_plan, api_key, instruction.strip(), fields)
