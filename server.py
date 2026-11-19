"""Serve the standalone Event Manager on this computer."""

import argparse
from contextlib import closing, contextmanager
from dataclasses import dataclass
from pathlib import Path
import secrets
import sqlite3

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.trustedhost import TrustedHostMiddleware
import uvicorn

from event_state import install


ROOT = Path(__file__).resolve().parent
DEFAULT_DATA = Path.home() / '.event-manager'
WORKSPACE_ID = 'standalone'
SESSION_COOKIE = 'event_manager_session'


@dataclass(frozen=True)
class Database:
    home: Path

    @contextmanager
    def connect(self):
        with closing(sqlite3.connect(self.home / 'workspace.sqlite3', timeout=10)) as db, db:
            yield db


@dataclass(frozen=True)
class Workspace:
    database: Database
    identity: str = WORKSPACE_ID


def create_app(data_dir=DEFAULT_DATA):
    home = Path(data_dir).expanduser().resolve()
    home.mkdir(mode=0o700, parents=True, exist_ok=True)
    workspace = Workspace(Database(home))
    session_token, csrf_token = secrets.token_urlsafe(32), secrets.token_urlsafe(32)
    app = FastAPI(title='Event Manager', docs_url=None, redoc_url=None, openapi_url=None)
    app.add_middleware(TrustedHostMiddleware, allowed_hosts=['127.0.0.1', 'localhost'])

    def require_session(request):
        supplied = request.cookies.get(SESSION_COOKIE, '').encode()
        if not secrets.compare_digest(supplied, session_token.encode()):
            raise HTTPException(403, 'Reload Event Manager to open a local session.')

    def pinned(request, write=False, expected=None):
        require_session(request)
        expected = request.headers.get('x-parcel-workspace') if expected is None else expected
        if expected != WORKSPACE_ID:
            raise HTTPException(409, 'The workspace changed. Reload Event Manager.')
        if write:
            origin = f'{request.url.scheme}://{request.url.netloc}'
            supplied = request.headers.get('x-chaos-csrf', '').encode()
            if (request.headers.get('origin') != origin or
                    not secrets.compare_digest(supplied, csrf_token.encode())):
                raise HTTPException(403, 'Reload Event Manager before saving.')
        return workspace

    @app.middleware('http')
    async def response_headers(request, call_next):
        response = await call_next(request)
        response.headers['Cache-Control'] = 'no-store'
        response.headers['Referrer-Policy'] = 'no-referrer'
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['Content-Security-Policy'] = (
            "default-src 'none'; script-src 'self'; style-src 'self' 'unsafe-inline'; "
            "connect-src 'self'; img-src 'self' data: blob:; font-src 'self'; "
            "base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
        )
        return response

    @app.get('/', include_in_schema=False)
    @app.get('/event-coordinator/', include_in_schema=False)
    def page():
        response = FileResponse(ROOT / 'app.html', media_type='text/html')
        response.set_cookie(SESSION_COOKIE, session_token, httponly=True, samesite='strict')
        return response

    @app.get('/api/session')
    def session(request: Request):
        require_session(request)
        return {'user': {'profile': WORKSPACE_ID}, 'csrf': csrf_token}

    install(app, pinned)
    app.mount('/static', StaticFiles(directory=ROOT / 'static'), name='static')
    return app


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=8789, help='Local HTTP port (default %(default)s).')
    parser.add_argument('--data-dir', type=Path, default=DEFAULT_DATA,
                        help='Directory for the encrypted workspace and its key (default %(default)s).')
    args = parser.parse_args()
    uvicorn.run(create_app(args.data_dir), host='127.0.0.1', port=args.port,
                proxy_headers=False, timeout_graceful_shutdown=2)


if __name__ == '__main__':
    main()
