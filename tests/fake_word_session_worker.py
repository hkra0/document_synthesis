# -*- coding: utf-8 -*-
"""替身会话 worker：真实的 ``serve_session`` 主循环 + 假 COM 对象。

供 ``tests/test_win_com_backend.py`` 以子进程方式驱动 :class:`WordSession`，不需要 Word。
行为由环境变量控制：

- ``FAKE_WORD_FAULT``：``<序号>:<故障>``，在该 Word 实例处理第 N 个文档的 Repaginate 时
  注入故障；故障为 ``hang``（挂起）、``crash``（worker 直接退出）或 ``com``（RPC 断连）。
- ``FAKE_WORD_QUIT_HANG``：为 1 时 Quit 一直不返回，用于验证 Quit 看门狗。
- ``FAKE_WORD_RECORD_PID``：把这个 PID 当作“本次启动的 Word”写入 pid 文件。
"""

import faulthandler
import json
import os
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from lib.office import win_com, win_com_worker as worker  # noqa: E402
from test_win_com_backend import FakeComError, FakeWordApp  # noqa: E402


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    fault = os.environ.get("FAKE_WORD_FAULT", "")
    fault_index, fault_kind = (int(fault.split(":")[0]), fault.split(":")[1]) if fault else (0, "")
    quit_hang = os.environ.get("FAKE_WORD_QUIT_HANG") == "1"
    record_pid = int(os.environ.get("FAKE_WORD_RECORD_PID") or 0)

    # 离线测试不读取真实进程表：不写启动前快照，只在指定时记录替身“Word”进程。
    worker._write_pending_record = lambda pid_file, image: None

    def record(pid_file, pid, image):
        if record_pid and pid_file:
            name, created = win_com.process_identity(record_pid)
            Path(pid_file).write_text(json.dumps({"pid": record_pid, "image": name, "created": created}),
                                      encoding="utf-8")

    worker._record_instance = record

    def on_repaginate(doc):
        if len(doc.app.documents) != fault_index:
            return
        if fault_kind == "hang":
            time.sleep(120)
        elif fault_kind == "crash":
            sys.stderr.flush()
            os._exit(3)
        elif fault_kind == "com":
            raise FakeComError(-2147023174, "The RPC server is unavailable.", None, None)

    class App(FakeWordApp):
        def Quit(self, SaveChanges=None):  # noqa: N802,N803
            if quit_hang:
                time.sleep(120)
            super().Quit(SaveChanges)

    def dispatch(progid):
        return App(
            bookmark_starts={"cover": 1, "body": 9},
            pages_by_start={1: (1, 1), 9: (3, 1)},
            on_repaginate=on_repaginate,
        )

    start = json.loads(sys.stdin.readline() or "{}")
    worker._completion.watchdog_enabled = True
    worker.QUIT_GRACE_SECONDS = float(os.environ.get("FAKE_WORD_QUIT_GRACE", "30"))
    return worker.serve_session(
        start,
        sys.stdin.readline,
        lambda text: print(text, flush=True),
        dispatch,
        lambda app: 0,
        arm_stack_dump=lambda seconds: faulthandler.dump_traceback_later(seconds, file=sys.stderr),
        cancel_stack_dump=faulthandler.cancel_dump_traceback_later,
    )


if __name__ == "__main__":
    sys.exit(main())
