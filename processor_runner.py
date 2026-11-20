"""Run manually requested text rules in a disposable, restricted CPython process."""

import asyncio
import ctypes
from functools import lru_cache
import json
import os
from pathlib import Path
import re
import selectors
import shutil
import stat
import subprocess
import sys
import tempfile
import threading
import time

from fastapi import HTTPException, Request


RUN_PATH = '/api/event-coordinator/processor/run'
SANDBOX = '/usr/bin/sandbox-exec'
EMPTY = '/private/var/empty'
APPLE_PYTHON = Path('/Library/Developer/CommandLineTools/Library/Frameworks/Python3.framework/Versions/3.9/bin/python3.9')
PYTHON_FLAGS = ['-I', '-S', '-B', '-X', 'utf8']
FIELDS = {'nickname', 'group', 'notes', 'sample'}
MAX_CELLS = 2000
MAX_VALUE = 8000
MAX_TOTAL = 250000
MAX_RULES = 20
MAX_SCRIPT = 16000
MAX_SCRIPT_TOTAL = 64000
MAX_APPLICATIONS = 10000
MAX_WIRE_BYTES = 4 * 1024 * 1024
MAX_MEMORY = 256 * 1024 * 1024
CPU_SECONDS = 2
WALL_SECONDS = 4
POLL_SECONDS = 0.025
READY = b'ready\n'
SLOTS = threading.BoundedSemaphore(2)

WORKER = f'''\
import contextlib, json, resource, sys

destination = sys.stdout.buffer
for name, maximum in ((resource.RLIMIT_CPU, {CPU_SECONDS}), (resource.RLIMIT_CORE, 0),
                      (resource.RLIMIT_FSIZE, 0), (resource.RLIMIT_NOFILE, 32)):
    resource.setrlimit(name, (maximum, maximum))
destination.write({READY!r})
destination.flush()
data = json.load(sys.stdin)

def run(data):
    functions = {{}}
    for index, rule in enumerate(data['rules'], 1):
        namespace = {{'__name__': '__processor_rule__'}}
        try:
            exec(compile(rule['script'], '<processor rule>', 'exec'), namespace)
        except BaseException as error:
            line = f' on line {{error.lineno}}' if isinstance(error, SyntaxError) else ''
            return {{'error': f'Rule {{index}} could not start{{line}} ({{type(error).__name__[:40]}}).'}}
        process = namespace.get('process')
        if not callable(process):
            return {{'error': f'Rule {{index}} must define process(value).'}}
        functions[rule['id']] = index, process

    result = []
    total = 0
    for position, cell in enumerate(data['cells'], 1):
        value = cell['value']
        for identity in cell['rules']:
            index, process = functions[identity]
            try:
                value = process(value)
            except BaseException as error:
                return {{'error': f'Rule {{index}} failed on selected cell {{position}} ({{type(error).__name__[:40]}}).'}}
            if not isinstance(value, str) or len(value) > {MAX_VALUE}:
                return {{'error': f'Rule {{index}} must return text of up to {MAX_VALUE:,} characters.'}}
        total += len(value)
        if total > {MAX_TOTAL}:
            return {{'error': 'The combined results are too long. Select fewer cells.'}}
        result.append({{'id': cell['id'], 'value': value}})
    return {{'cells': result}}

with contextlib.redirect_stdout(sys.stderr):
    result = run(data)
destination.write(json.dumps(result).encode('ascii'))
destination.flush()
'''


class RUsage(ctypes.Structure):
    # macOS sys/resource.h, rusage_info_v0.
    _fields_ = [('uuid', ctypes.c_uint8 * 16)] + [
        (name, ctypes.c_uint64) for name in ('user_time', 'system_time', 'package_wakeups', 'interrupt_wakeups',
                                             'pageins', 'wired', 'resident', 'footprint', 'start_time', 'exit_time')
    ]


def unavailable():
    return HTTPException(503, 'The local Python sandbox is unavailable. No script was started.')


def apple_python():
    try:
        if APPLE_PYTHON.resolve(strict=True) == APPLE_PYTHON and all(
                metadata.st_uid == 0 and not metadata.st_mode & (stat.S_IWGRP | stat.S_IWOTH)
                for path in (APPLE_PYTHON, *APPLE_PYTHON.parents) for metadata in (path.stat(),)):
            return str(APPLE_PYTHON)
    except OSError:
        pass
    return None


@lru_cache(maxsize=1)
def runtime():
    if sys.platform != 'darwin' or not os.access(SANDBOX, os.X_OK):
        raise unavailable()
    path = os.environ.get('EVENT_COORDINATOR_PROCESSOR_PYTHON')
    if not path:
        path = apple_python() if getattr(sys, 'frozen', False) else getattr(sys, '_base_executable', sys.executable)
    if not path or not Path(path).is_absolute():
        raise unavailable()
    try:
        executable = Path(path).resolve(strict=True)
        probe = subprocess.run([str(executable), *PYTHON_FLAGS, '-c', (
            'import json, sys, sysconfig; print(json.dumps([sys.implementation.name, '
            'sysconfig.get_path("stdlib"), sys.platlibdir, sysconfig.get_config_var("LDLIBRARY"), '
            'sysconfig.get_config_var("PYTHONFRAMEWORK")]))'
        )], cwd=EMPTY, env={}, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
            timeout=2, check=True)
        implementation, library, libdir, dylib, framework = json.loads(probe.stdout)
        stdlib = Path(library).resolve(strict=True)
        if (implementation != 'cpython' or not re.fullmatch(r'python3\.\d+', stdlib.name) or
                not (stdlib / 'encodings').is_dir() or not re.fullmatch(r'lib(?:64)?', libdir)):
            raise ValueError('Unsupported Python runtime')
        if framework in ('Python', 'Python3'):
            # Framework bin/python is a spawning stub. Launch the interpreter itself.
            executable = (stdlib.parent.parent / 'Resources/Python.app/Contents/MacOS/Python').resolve(strict=True)
        executable.relative_to(stdlib.parent.parent)
        libproc = ctypes.CDLL('/usr/lib/libproc.dylib')
        measure = libproc.proc_pid_rusage
        measure.argtypes = (ctypes.c_int, ctypes.c_int, ctypes.POINTER(RUsage))
        measure.restype = ctypes.c_int
    except (OSError, subprocess.SubprocessError, ValueError, TypeError, AttributeError):
        raise unavailable() from None
    return executable, stdlib, libdir, dylib, framework, measure


def sandbox_profile(launcher, stdlib, lib_alias, dylib, framework_files):
    def quoted(path):
        return json.dumps(str(path), ensure_ascii=False)

    root = lib_alias.parent
    entries = [launcher, *launcher.parents[:len(launcher.parts) - len(root.parts)], lib_alias, *framework_files]
    if dylib and re.fullmatch(r'libpython3\.\d+\.dylib', dylib):
        entries.extend((stdlib.parent / dylib, lib_alias / dylib))
    for prefix in (stdlib.parent, lib_alias):
        zip_name = stdlib.name.replace('.', '') + '.zip'
        entries.append(prefix / zip_name)
    literals = ' '.join(f'(literal {quoted(path)})' for path in entries)
    library_roots = {stdlib, lib_alias / stdlib.name}
    libraries = '\n'.join(
        f'(allow file-read* (require-all (subpath {quoted(root)}) '
        f'(require-not (subpath {quoted(root / "site-packages")})) '
        f'(require-not (subpath {quoted(root / "dist-packages")}))))'
        for root in sorted(library_roots)
    )
    return f'''(version 1)
(deny default)
(allow process-exec (literal {quoted(launcher)}))
(allow file-read* {literals})
{libraries}
(allow file-read* (subpath "/usr/lib") (subpath "/System/Library/Frameworks") (subpath "/System/Library/PrivateFrameworks"))
(allow file-read-data (literal "/"))
(allow file-read* (literal "/dev/random") (literal "/dev/urandom"))
'''


def read_result(process, launcher, payload, measure):
    deadline = time.monotonic() + WALL_SECONDS
    reply = bytearray()
    payload = memoryview(payload)
    ready = False
    stats = RUsage()
    with selectors.DefaultSelector() as selector:
        stdout, stdin = process.stdout.fileno(), process.stdin.fileno()
        os.set_blocking(stdout, False)
        os.set_blocking(stdin, False)
        selector.register(stdout, selectors.EVENT_READ)
        while selector.get_map() or process.poll() is None:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise HTTPException(422, 'The Python rule ran too long. Try fewer cells or simplify the rule.')
            if process.poll() is None:
                if measure(process.pid, 0, ctypes.byref(stats)) != 0 and process.poll() is None:
                    raise unavailable()
                if max(stats.resident, stats.footprint) > MAX_MEMORY:
                    raise HTTPException(422, 'The Python rule used too much memory.')
            delay = min(POLL_SECONDS, remaining)
            if not selector.get_map():
                time.sleep(delay)
                continue
            for event, _ in selector.select(delay):
                if event.fd == stdin:
                    try:
                        payload = payload[os.write(stdin, payload[:65536]):]
                    except BrokenPipeError:
                        payload = payload[len(payload):]
                    if not payload:
                        selector.unregister(stdin)
                        process.stdin.close()
                    continue
                chunk = os.read(stdout, 65536)
                if not chunk:
                    selector.unregister(stdout)
                    continue
                reply.extend(chunk)
                if len(reply) > MAX_WIRE_BYTES:
                    raise HTTPException(422, 'The Python rule produced too much output.')
                if not ready and len(reply) >= len(READY):
                    if reply != READY:
                        raise unavailable()
                    # After startup there is no longer a path the sandbox permits executing.
                    launcher.unlink()
                    reply.clear()
                    ready = True
                    selector.register(stdin, selectors.EVENT_WRITE)
    if not ready:
        raise unavailable()
    if process.returncode != 0:
        raise HTTPException(422, 'The Python rule stopped before returning a result.')
    return reply


def sandboxed(data):
    executable, stdlib, libdir, dylib, framework, measure = runtime()
    payload = json.dumps(data, separators=(',', ':')).encode()
    with tempfile.TemporaryDirectory(prefix='event-manager-processor-') as directory:
        root = Path(directory).resolve()
        launcher = root / executable.relative_to(stdlib.parent.parent)
        launcher.parent.mkdir(parents=True, exist_ok=True)
        lib_alias = root / libdir
        lib_alias.symlink_to(stdlib.parent, target_is_directory=True)
        framework_files = []
        if framework in ('Python', 'Python3'):
            library = (stdlib.parent.parent / framework).resolve(strict=True)
            framework_alias = root / framework
            framework_alias.symlink_to(library)
            framework_files = [library, framework_alias]
        try:
            os.link(executable, launcher)
        except OSError:
            shutil.copy2(executable, launcher)
        profile = sandbox_profile(launcher, stdlib, lib_alias, dylib, framework_files)
        try:
            process = subprocess.Popen([SANDBOX, '-p', profile, str(launcher), *PYTHON_FLAGS, '-c', WORKER],
                                       cwd=EMPTY, env={}, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                       stderr=subprocess.DEVNULL, close_fds=True, start_new_session=True)
        except OSError:
            raise unavailable() from None
        with process:
            try:
                return read_result(process, launcher, payload, measure)
            finally:
                if process.poll() is None:
                    process.kill()


def text(value, limit):
    return isinstance(value, str) and len(value) <= limit and not re.search(r'[\ud800-\udfff]', value)


def validate(body):
    rules, cells = body.get('rules'), body.get('cells')
    if (set(body) != {'rules', 'cells'} or not isinstance(rules, list) or not 1 <= len(rules) <= MAX_RULES or
            not isinstance(cells, list) or not 1 <= len(cells) <= MAX_CELLS):
        raise HTTPException(422, f'Choose up to {MAX_RULES} Python rules and {MAX_CELLS:,} text cells.')
    rule_ids = set()
    script_length = 0
    for rule in rules:
        if (not isinstance(rule, dict) or set(rule) != {'id', 'script'} or
                not text(rule['id'], 128) or not rule['id'] or rule['id'] in rule_ids or
                not text(rule['script'], MAX_SCRIPT) or not rule['script'].strip()):
            raise HTTPException(422, f'Each rule needs an ID and a script of up to {MAX_SCRIPT:,} characters.')
        rule_ids.add(rule['id'])
        script_length += len(rule['script'])
    cell_ids = set()
    value_length = applications = 0
    for cell in cells:
        if (not isinstance(cell, dict) or set(cell) != {'id', 'field', 'value', 'rules'} or
                not text(cell['id'], 128) or not cell['id'] or cell['id'] in cell_ids or
                not isinstance(cell['field'], str) or cell['field'] not in FIELDS or not text(cell['value'], MAX_VALUE) or
                not isinstance(cell['rules'], list) or not 1 <= len(cell['rules']) <= MAX_RULES or
                any(not isinstance(rule, str) or rule not in rule_ids for rule in cell['rules']) or
                len(set(cell['rules'])) != len(cell['rules'])):
            raise HTTPException(422, f'Choose valid nickname, group, notes or sample text cells of up to {MAX_VALUE:,} characters.')
        cell_ids.add(cell['id'])
        value_length += len(cell['value'])
        applications += len(cell['rules'])
    if script_length > MAX_SCRIPT_TOTAL or value_length > MAX_TOTAL or applications > MAX_APPLICATIONS:
        raise HTTPException(422, 'The combined selection and rules are too large. Select fewer cells or rules.')
    return rules, cells


def run(rules, cells):
    # Keep sheet metadata and caller IDs outside the sandbox regardless of the rule text.
    packet = {'rules': rules, 'cells': [{'id': index, 'value': cell['value'], 'rules': cell['rules']}
                                       for index, cell in enumerate(cells)]}
    try:
        output = json.loads(sandboxed(packet))
    except (ValueError, RecursionError):
        raise HTTPException(422, 'The Python rule did not return a valid result.') from None
    except OSError:
        raise unavailable() from None
    if isinstance(output, dict) and set(output) == {'error'} and text(output['error'], 500):
        raise HTTPException(422, output['error'])
    if isinstance(output, dict) and set(output) == {'cells'} and isinstance(output['cells'], list) and len(output['cells']) == len(cells):
        result = []
        for index, (cell, returned) in enumerate(zip(cells, output['cells'])):
            if (not isinstance(returned, dict) or set(returned) != {'id', 'value'} or
                    type(returned['id']) is not int or returned['id'] != index or not text(returned['value'], MAX_VALUE)):
                break
            result.append({'id': cell['id'], 'value': returned['value']})
        else:
            if sum(len(cell['value']) for cell in result) <= MAX_TOTAL:
                return {'cells': result}
    raise HTTPException(422, 'The Python rule did not return one text value for every selected cell.')


def install_processor(app, pinned):
    from event_state import read_json

    @app.post(RUN_PATH)
    async def run_processor(request: Request):
        pinned(request, True)
        body = await read_json(request, limit=MAX_WIRE_BYTES)
        rules, cells = validate(body)
        if not SLOTS.acquire(blocking=False):
            raise HTTPException(429, 'The Python processor is busy. Try again shortly.')

        def run_and_release():
            try:
                return run(rules, cells)
            finally:
                SLOTS.release()

        # A cancelled browser request should not free a slot while its worker still runs.
        return await asyncio.shield(asyncio.to_thread(run_and_release))
