import threading
import app

def run_flask():
    app.app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)

if __name__ == "__main__":
    t = threading.Thread(target=run_flask)
    t.daemon = True
    t.start()
    print("Flask started in background thread.")
    import time
    while True:
        time.sleep(1)

