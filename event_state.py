"""Private, revisioned Event Coordinator state. The DB contains ciphertext only."""
import asyncio
import json
import os
import re
from functools import lru_cache
from pathlib import Path

from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from fastapi import HTTPException, Request
from fastapi.responses import JSONResponse, StreamingResponse

KEY_FILE = 'event-coordinator.key'
CARD_FIELDS = {'number', 'month', 'year', 'cvv', 'holder', 'brand', 'paymentType', 'firstName', 'lastName',
               'street', 'street2', 'city', 'state', 'zip', 'country'}
STATE_FIELDS = {'version', 'people', 'addresses', 'accounts', 'cards', 'groups', 'events', 'selected',
                'serials', 'accountKeys', 'sourceBaselines', 'ledger', 'accountGroups', 'removedAccountKeys',
                'labelExpressions'}
BLOCKED = {'cvv', 'cvv2', 'cvc', 'cvc2', 'ccv', 'csc', 'cid', 'securitycode', 'cardsecuritycode',
           'cccvv', 'cccvc', 'cardcvv', 'cardcvc', 'paymentsecuritycode', 'cardverificationcode',
           'cardverificationvalue', 'password', 'pan', 'cardnumber', 'number', 'proto', 'constructor', 'prototype'}
LIMIT = 12 * 1024 * 1024


def reject():
    raise HTTPException(400, 'Invalid coordinator data.')


def valid_label_expressions(expressions):
    if not isinstance(expressions, list) or len(expressions) > 100:
        return False
    ids = set()
    for item in expressions:
        if (not isinstance(item, dict) or not {'id', 'name', 'template'} <= item.keys()
                or item.keys() - {'id', 'name', 'template', 'field', 'enabled', 'language'}):
            return False
        if not all(isinstance(item[key], str) and item[key].strip() for key in ('id', 'name', 'template')):
            return False
        if item.get('field', 'nickname') not in ('nickname', 'group', 'notes') or type(item.get('enabled', True)) is not bool:
            return False
        key = item['id']
        if not re.fullmatch(r'[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}', key) or key in ids:
            return False
        if 'language' in item and item['language'] != 'python':
            return False
        script_limit = 16000 if item.get('language') == 'python' else 400
        if len(item['name']) > 80 or len(item['template']) > script_limit:
            return False
        ids.add(key)
    return True


def inspect_tree(value, depth=0):
    if depth > 20:
        reject()
    if isinstance(value, dict):
        for key, child in value.items():
            if re.sub(r'[^a-z0-9]', '', key.lower()) in BLOCKED or len(key) > 512:
                reject()
            inspect_tree(child, depth + 1)
    elif isinstance(value, list):
        if len(value) > 100000:
            reject()
        for child in value:
            inspect_tree(child, depth + 1)


def valid_card_number(digits):
    if not re.fullmatch(r'[0-9]{12,19}', digits):
        return False
    total = 0
    for index, character in enumerate(reversed(digits)):
        digit = int(character)
        if index % 2:
            digit *= 2
            if digit > 9:
                digit -= 9
        total += digit
    return total > 0 and total % 10 == 0


def valid_list_config(config):
    if not isinstance(config, dict) or not isinstance(config.get('nicknameCounters'), dict):
        return False
    if any(not isinstance(config.get(key), list) for key in ('chosen','cards','names','edits','rows')):
        return False
    if any(not isinstance(item,list) or len(item)!=2 for item in config['names']+config['edits']):
        return False
    return all(isinstance(row,list) and len(row)==3 and type(row[0]) is int for row in config['rows'])


def valid_layout(state):
    selected, serials = state.get('selected'), state.get('serials')
    if not isinstance(selected, dict) or not isinstance(serials, dict):
        return False
    legacy = state['version'] == 1
    if legacy and selected.get('bot') not in ('sniped', 'shikari'):
        return False
    active = selected.get('event')
    if active is not None and not isinstance(active, str):
        return False
    if any(type(serials.get(field)) != int for field in ('sourceSerial', 'eventSerial', 'localAccountSerial', 'accountSerial', 'nextCardId')):
        return False
    ids = set()
    for event in state['events']:
        if not isinstance(event, dict) or not isinstance(event.get('id'), str) or not re.fullmatch(r'event-[1-9][0-9]*', event['id']) or event['id'] in ids:
            return False
        ids.add(event['id'])
        if legacy:
            if not all(valid_list_config(event.get(bot)) for bot in ('sniped', 'shikari')):
                return False
            continue
        if not isinstance(event.get('name'), str) or not isinstance(event.get('lists'), list):
            return False
        list_ids = set()
        for config in event['lists']:
            if not valid_list_config(config):
                return False
            identity = config.get('id')
            if not isinstance(identity, str) or not re.fullmatch(r'(?:list-[0-9a-f]{32}|sniped|shikari)', identity) or identity in list_ids:
                return False
            if config.get('format') not in ('sniped', 'shikari', 'aycd') or not isinstance(config.get('name'), str) or not config['name'].strip() or len(config['name']) > 200:
                return False
            list_ids.add(identity)
        selected_list = event.get('selectedListId')
        if list_ids:
            if not isinstance(selected_list, str) or selected_list not in list_ids:
                return False
        elif selected_list is not None:
            return False
    return active in ids if ids else not legacy and active is None


def read_key(home, create=False):
    path = Path(home) / KEY_FILE
    if create and not path.exists():
        key = AESGCM.generate_key(bit_length=256)
        try:
            fd = os.open(path, os.O_CREAT | os.O_WRONLY | os.O_EXCL, 0o600)
        except FileExistsError:
            pass
        else:
            with os.fdopen(fd, 'wb') as output:
                output.write(key)
                output.flush()
                os.fsync(output.fileno())
    try:
        key = path.read_bytes()
        if len(key) == 32:
            return key
    except OSError:
        pass
    raise HTTPException(503, 'Restore the coordinator encryption key before continuing.')


def context_bytes(resources, revision):
    return f'event-coordinator/v1:{resources.identity}:{revision}'.encode()


def coordinator_table_exists(db):
    return db.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='event_coordinator_state'").fetchone() is not None


def read_revision(resources):
    with resources.database.connect() as db:
        if coordinator_table_exists(db):
            row = db.execute('SELECT revision FROM event_coordinator_state WHERE slot=1').fetchone()
            if row:
                return row[0]
    return 0


def read_box(db, resources):
    row = db.execute('SELECT revision, value FROM event_coordinator_state WHERE slot=1').fetchone() if coordinator_table_exists(db) else None
    if not row:
        return 0, {'state': None, 'cards': {}, 'passwords': {}}
    revision, blob = row
    key = read_key(resources.database.home)
    try:
        return revision, json.loads(AESGCM(key).decrypt(blob[:12], blob[12:], context_bytes(resources, revision)))
    except Exception:
        raise HTTPException(503, 'Coordinator state could not be opened. Restore the matching key and backup.') from None


@lru_cache(maxsize=1)
def network_patterns():
    source = Path(__file__).parent / 'static' / 'event-coordinator' / 'card-networks.json'
    return [(network, re.compile(pattern)) for network, pattern in json.loads(source.read_text()).items()]


def populate_card_metadata(box):
    state = box['state']
    for account in (state or {}).get('accounts', []):
        account['hasPassword'] = bool(box['passwords'].get(str(account['id'])))
    for record in (state or {}).get('cards', []):
        secret = box['cards'].get(record['id'], {})
        record['hasSecurityCode'] = bool(secret.get('cvv'))
        brand = secret.get('brand') if secret.get('number') else secret.get('brand') or record.get('network')
        if isinstance(brand, str) and brand.strip() and len(brand) <= 80 and not re.search(r'(?:[0-9][ -]?){12}', brand):
            record['network'] = brand.strip()
        elif secret.get('number'):
            record['network'] = next((network for network, pattern in network_patterns()
                                      if pattern.fullmatch(secret['number'])), '')
        else:
            record.pop('network', None)


def secret_revisions(box, revision):
    if 'secretRevisions' in box:
        return box['secretRevisions']
    # A legacy box proves only that these credentials existed at this revision.
    return {kind: dict.fromkeys(box[kind], revision) for kind in ('cards', 'passwords')}


def save_box(resources, revision, state, changes):
    with resources.database.connect() as db:
        db.execute('BEGIN IMMEDIATE')
        current, box = read_box(db, resources)
        if revision != current:
            raise HTTPException(409, 'Another session saved changes. Reload the current revision.')
        if not isinstance(state, dict) or type(state.get('version')) is not int or state['version'] not in (1,2) or set(state) - STATE_FIELDS:
            reject()
        if state['version'] == 1 and (box.get('state') or {}).get('version') == 2:
            raise HTTPException(409, 'This workspace now uses named lists. Refresh Event Manager.')
        for field in ('people', 'addresses', 'accounts', 'cards', 'groups', 'events', 'sourceBaselines', 'accountKeys'):
            if not isinstance(state.get(field), list):
                reject()
        groups = state.get('accountGroups', [])
        if not isinstance(groups, list) or any(not isinstance(group, dict) or
                any(not isinstance(group.get(key), str) or not group[key] for key in ('id', 'retailer', 'name')) for group in groups):
            reject()
        removed = state.get('removedAccountKeys', [])
        if not isinstance(removed, list) or any(not isinstance(key, str) for key in removed):
            reject()
        inspect_tree(state)
        if (not valid_layout(state) or not valid_label_expressions(state.get('labelExpressions', []))
                or not isinstance(changes, dict) or set(changes)-{'cards', 'passwords'}):
            reject()
        if not all(isinstance(row, dict) and isinstance(row.get('id'), int) for row in state['accounts']):
            reject()
        if not all(isinstance(row, dict) and isinstance(row.get('id'), str) for row in state['cards']):
            reject()
        card_ids = {row['id'] for row in state['cards']}
        login_ids = {str(row['id']) for row in state['accounts']}
        card_updates, passwords = changes.get('cards', {}), changes.get('passwords', {})
        if not isinstance(card_updates, dict) or not isinstance(passwords, dict):
            reject()
        if set(card_updates)-card_ids or set(passwords)-login_ids:
            reject()
        for card_id, card in card_updates.items():
            if not isinstance(card, dict) or set(card)-CARD_FIELDS or any(not isinstance(v, str) or len(v) > 500 for v in card.values()):
                reject()
            number = card.get('number')
            if number and number != box['cards'].get(card_id, {}).get('number') and not valid_card_number(number):
                reject()
            if card.get('month') and not re.fullmatch(r'0[1-9]|1[0-2]', card['month']):
                reject()
            if card.get('year') and not re.fullmatch(r'[0-9]{4}', card['year']):
                reject()
            if card.get('cvv') and not re.fullmatch(r'[0-9]{3,4}', card['cvv']):
                reject()
        if any(not isinstance(password, str) or len(password) > 4096 for password in passwords.values()):
            reject()
        versions = secret_revisions(box, current)
        box['state'] = state
        revision += 1
        for card_id, patch in card_updates.items():
            previous = box['cards'].get(card_id, {})
            secret = previous | patch
            if secret.get('cvv') == '':
                del secret['cvv']
            if secret != previous:
                versions['cards'][card_id] = revision
            box['cards'][card_id] = secret
        for account_id, password in passwords.items():
            if password != box['passwords'].get(account_id, ''):
                versions['passwords'][account_id] = revision
        box['cards'] = {key: value for key, value in box['cards'].items() if key in card_ids}
        box['passwords'] = {key: value for key, value in (box['passwords'] | passwords).items() if key in login_ids and value}
        box['secretRevisions'] = {
            'cards': {key: value for key, value in versions['cards'].items() if key in card_ids},
            'passwords': {key: value for key, value in versions['passwords'].items() if key in login_ids}}
        raw = json.dumps(box, ensure_ascii=False, separators=(',', ':'), allow_nan=False).encode()
        if len(raw) > LIMIT:
            raise HTTPException(413, 'Coordinator state is too large.')
        key = read_key(resources.database.home, create=True)
        nonce = os.urandom(12)
        encoded = nonce + AESGCM(key).encrypt(nonce, raw, context_bytes(resources, revision))
        db.execute('CREATE TABLE IF NOT EXISTS event_coordinator_state (slot INTEGER PRIMARY KEY CHECK(slot=1), revision INTEGER NOT NULL, value BLOB NOT NULL)')
        db.execute('INSERT OR REPLACE INTO event_coordinator_state(slot,revision,value) VALUES(1,?,?)', (revision, encoded))
    return revision, box['secretRevisions']


async def read_json(request, limit=LIMIT):
    if request.headers.get('content-type', '').split(';')[0] != 'application/json':
        raise HTTPException(415, 'JSON required.')
    raw = bytearray()
    async for chunk in request.stream():
        raw.extend(chunk)
        if len(raw) > limit:
            raise HTTPException(413, 'Coordinator request is too large.')
    try:
        value = json.loads(raw)
        if isinstance(value, dict):
            return value
    except (ValueError, RecursionError):
        pass
    reject()


def install(app, pinned):
    listeners = {}

    from event_ai import install as install_ai
    from processor_runner import install_processor
    install_ai(app, pinned)
    install_processor(app, pinned)

    @app.get('/api/event-coordinator/revision')
    def current_revision(request: Request):
        resources = pinned(request)
        return JSONResponse({'profileId': resources.identity, 'revision': read_revision(resources)},
                            headers={'Cache-Control': 'private, no-store'})

    @app.get('/api/event-coordinator/events')
    async def events(request: Request, workspace: str):
        # Native EventSource uses cookies and cannot supply custom request headers.
        resources = pinned(request, expected=workspace)

        async def notices():
            wake = asyncio.Event()
            subscribers = listeners.setdefault(workspace, set())
            subscribers.add(wake)
            previous = None
            try:
                while not getattr(app.state, 'live_stopping', False) and not await request.is_disconnected():
                    wake.clear()
                    try:
                        pinned(request, expected=workspace)
                    except HTTPException:
                        yield 'event: login\ndata: {}\n\n'
                        return
                    current = await asyncio.to_thread(read_revision, resources)
                    if current != previous:
                        yield 'data: ' + json.dumps({'profileId': workspace, 'revision': current}) + '\n\n'
                        previous = current
                    else:
                        yield ': keepalive\n\n'
                    try:
                        await asyncio.wait_for(wake.wait(), timeout=15)
                    except asyncio.TimeoutError:
                        # Recheck authentication and detect writes by another worker.
                        pass
            finally:
                subscribers.discard(wake)
                if not subscribers:
                    listeners.pop(workspace, None)

        return StreamingResponse(notices(), media_type='text/event-stream',
                                 headers={'Cache-Control': 'no-store', 'X-Accel-Buffering': 'no'})

    @app.get('/api/event-coordinator/state')
    def read(request: Request):
        resources = pinned(request)
        with resources.database.connect() as db:
            revision, box = read_box(db, resources)
        populate_card_metadata(box)
        return {'profileId': resources.identity, 'revision': revision, 'state': box['state'],
                'secretRevisions': secret_revisions(box, revision)}

    @app.put('/api/event-coordinator/state')
    async def write(request: Request):
        resources = pinned(request, True)
        body = await read_json(request)
        if set(body) != {'revision', 'state', 'changes'} or type(body['revision']) != int:
            reject()
        revision, versions = save_box(resources, body['revision'], body['state'], body['changes'])
        for wake in listeners.get(resources.identity, ()):
            wake.set()
        return {'profileId': resources.identity, 'revision': revision, 'secretRevisions': versions}

    @app.post('/api/event-coordinator/secrets')
    async def selected_secrets(request: Request):
        resources = pinned(request, True)
        body = await read_json(request)
        if set(body) - {'cards', 'accounts', 'revision'} or type(body.get('revision')) != int:
            reject()
        for field in ('cards', 'accounts'):
            ids = body.get(field, [])
            if not isinstance(ids, list) or len(ids) > 100000 or any(not isinstance(item, str) or len(item) > 500 for item in ids):
                reject()
        with resources.database.connect() as db:
            revision, box = read_box(db, resources)
        if revision != body['revision']:
            raise HTTPException(409, 'The saved credentials changed. Reload the current revision.')
        return JSONResponse({'cards': {key: box['cards'][key] for key in body.get('cards', []) if key in box['cards']},
                             'passwords': {key: box['passwords'][key] for key in body.get('accounts', []) if key in box['passwords']}})
