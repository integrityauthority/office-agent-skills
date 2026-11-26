"""office_kernel's COM thread keeps answering COM while it waits for work."""
import sys
import threading
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / ".agents" / "skills" / "office-live" / "scripts"))
import office_kernel  # noqa: E402


class FakePythoncom:
    def __init__(self):
        self.pumps = 0

    def PumpWaitingMessages(self):
        self.pumps += 1


def test_waiting_for_a_call_pumps_messages():
    """Between calls the thread receives Excel's SheetChange events; without a
    pump Excel shows "waiting for another application to complete an OLE
    action" on every edit the user makes."""
    fake = FakePythoncom()
    request = ("print(1)", None)
    threading.Timer(0.3, lambda: office_kernel._requests.put(request)).start()
    assert office_kernel._next_request(fake) == request
    assert fake.pumps >= 3
