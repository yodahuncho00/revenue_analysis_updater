from pathlib import Path

from scripts import refresh_backoffice_session


def test_preflight_closes_browser_before_waiting_for_etls(tmp_path, monkeypatch):
    events = []
    config_path = tmp_path / "config.yaml"
    config_path.write_text(
        """
paths:
  storage_state_path: data/storage_state.json
backoffice:
  routes:
    dashboard: https://backoffice.example/dashboard
  browser:
    login_timeout_ms: 1000
pos_portal:
  payments_url: https://pos.example/payments
""",
        encoding="utf-8",
    )
    monkeypatch.setattr(refresh_backoffice_session, "PROJECT_ROOT", tmp_path)
    monkeypatch.setattr(refresh_backoffice_session, "_backoffice_ready", lambda *_args: True)
    monkeypatch.setattr(refresh_backoffice_session, "_pos_ready", lambda *_args: True)

    class Page:
        def goto(self, *_args, **_kwargs):
            pass

        def bring_to_front(self):
            pass

        def is_closed(self):
            return False

        def close(self):
            pass

    class Context:
        pages = [Page(), Page()]

        def storage_state(self, *, path):
            Path(path).write_text("{}", encoding="utf-8")

        def close(self):
            events.append("browser-closed")

    context = Context()

    class Extractor:
        def __init__(self, _config):
            pass

        def launch_persistent_context(self, *_args, **_kwargs):
            return context

    monkeypatch.setattr(refresh_backoffice_session, "PlaywrightExtractor", Extractor)

    class PlaywrightManager:
        def __enter__(self):
            return object()

        def __exit__(self, *_args):
            pass

    monkeypatch.setattr(
        refresh_backoffice_session, "sync_playwright", lambda: PlaywrightManager()
    )

    class Process:
        def __init__(self, source):
            self.source = source

        def wait(self):
            events.append(f"{self.source}-etl-finished")
            return 0

    etl_state_paths = []

    def check_state_snapshot(command, env):
        source = command[-1]
        state_path = Path(env["REVENUE_ETL_STORAGE_STATE"])
        assert state_path.is_file()
        etl_state_paths.append(state_path)
        return source

    def start_etl(command, *, env, **_kwargs):
        source = check_state_snapshot(command, env)
        events.append(f"{source}-etl-started")
        return Process(source)

    def run_etl(command, *, env, **_kwargs):
        source = check_state_snapshot(command, env)
        events.append(f"{source}-etl-started")
        events.append(f"{source}-etl-finished")
        return type("Result", (), {"returncode": 0})()

    monkeypatch.setattr(refresh_backoffice_session.subprocess, "Popen", start_etl)
    monkeypatch.setattr(refresh_backoffice_session.subprocess, "run", run_etl)

    assert refresh_backoffice_session.main(preflight=True) == 0
    assert events.index("browser-closed") < events.index("backoffice-etl-finished")
    assert events.index("browser-closed") < events.index("pos-etl-finished")
    assert events.index("backoffice-etl-finished") < events.index("pos-etl-started")
    assert all(not path.exists() for path in etl_state_paths)
