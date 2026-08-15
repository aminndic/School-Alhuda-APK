import os
import threading

_server_started = False
_lock = threading.Lock()


def start_server(data_dir):
    global _server_started
    with _lock:
        if _server_started:
            return
        os.environ['SCHOOL_DATA_DIR'] = data_dir
        # Import only after the writable Android data directory is known.
        import flask_app

        def run():
            flask_app.app.run(
                host='127.0.0.1',
                port=5000,
                debug=False,
                use_reloader=False,
                threaded=True,
            )

        t = threading.Thread(target=run, daemon=True, name='FlaskServer')
        t.start()
        _server_started = True
