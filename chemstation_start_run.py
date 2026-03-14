"""
chemstation_start_run.py

用途：
    通过 DDE 连接 Agilent ChemStation，
    触发一个固定的 macro（start_single_run.mac），
    从而启动一次单次 LC run。

作者建议：
    先在 ChemStation 内部手动运行一次 start_single_run.mac，
    确认 macro 本身没有问题，再用本脚本外部触发。

依赖：
    pip install pywin32

注意：
    1. 老版本 ChemStation 往往更适合 32-bit Python。
    2. ChemStation 的 DDE application name 需要你自己确认。
    3. macro 执行命令字符串在不同环境可能略有差异，
       下面给了几个常见写法可切换测试。
"""

import time
import logging
from typing import Optional

import dde


# =========================
# 用户需要修改的配置区
# =========================

# 1) ChemStation 的 DDE application name
#    你需要替换成你机器上的实际值。
#    之前你提到过可通过 _DDENAME$ 或系统配置获取。
CHEMSTATION_APP = "HPCORE"

# 2) 宏文件路径
MACRO_PATH = r"C:\Chem32\1\Macros\start_single_run.mac"

# 3) DDE topic
#    根据你给的 DDE 文档：
#    - SYSTEM：拿状态
#    - CPWAIT：同步执行命令
#    - CPNOWAIT：异步执行命令
TOPIC_SYSTEM = "SYSTEM"
TOPIC_EXECUTE = "CPWAIT"   # 推荐先用同步，更稳

# 4) 轮询超时参数
READY_TIMEOUT_S = 120
START_TIMEOUT_S = 60
POLL_INTERVAL_S = 1.0


# =========================
# 日志设置
# =========================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)


class ChemStationDDE:
    """
    对 ChemStation DDE 通信做一个简单封装。
    """

    def __init__(self, app_name: str):
        self.app_name = app_name
        self.server = dde.CreateServer()
        self.server.Create("python_chemstation_client")

    def _connect(self, topic: str):
        """
        建立到指定 topic 的 DDE 会话。
        """
        conv = dde.CreateConversation(self.server)
        conv.ConnectTo(self.app_name, topic)
        return conv

    def request(self, topic: str, item: str) -> str:
        """
        向指定 topic/item 发起 DDERequest。
        常用于读取 SYSTEM/Status 等状态信息。
        """
        conv = self._connect(topic)
        try:
            result = conv.Request(item)
            if isinstance(result, bytes):
                result = result.decode(errors="ignore")
            return str(result).strip()
        finally:
            conv.Disconnect()

    def execute(self, topic: str, command: str) -> None:
        """
        向指定 topic 发送 DDEExecute 命令。
        """
        conv = self._connect(topic)
        try:
            logger.info("DDE Execute -> topic=%s | command=%s", topic, command)
            conv.Exec(command)
        finally:
            conv.Disconnect()

    def get_status(self) -> str:
        """
        读取 ChemStation SYSTEM/Status。
        你给的 DDE 文档里说明会返回 READY 或 BUSY。
        """
        return self.request(TOPIC_SYSTEM, "Status")


def wait_until_ready(cs: ChemStationDDE, timeout_s: int = READY_TIMEOUT_S) -> None:
    """
    等待 ChemStation 进入 READY 状态。
    """
    logger.info("Waiting for ChemStation to become READY ...")
    t0 = time.time()

    while True:
        status = cs.get_status().upper()
        logger.info("Current status: %s", status)

        if "READY" in status:
            logger.info("ChemStation is READY.")
            return

        if time.time() - t0 > timeout_s:
            raise TimeoutError(f"ChemStation did not become READY within {timeout_s} s. Last status={status}")

        time.sleep(POLL_INTERVAL_S)


def wait_until_not_ready(cs: ChemStationDDE, timeout_s: int = START_TIMEOUT_S) -> None:
    """
    在触发 macro 后，等待状态从 READY 变成别的状态，
    作为“run 已经开始/系统已进入忙碌态”的一个简化判断。
    """
    logger.info("Waiting for ChemStation to leave READY state ...")
    t0 = time.time()

    while True:
        status = cs.get_status().upper()
        logger.info("Current status after trigger: %s", status)

        if "READY" not in status:
            logger.info("ChemStation left READY state. Run may have started.")
            return

        if time.time() - t0 > timeout_s:
            raise TimeoutError(
                f"ChemStation stayed READY for more than {timeout_s} s after trigger. "
                f"Macro may not have started correctly."
            )

        time.sleep(POLL_INTERVAL_S)


def build_macro_command(macro_path: str, mode: str = "style1") -> str:
    """
    构造触发 macro 的命令字符串。

    不同 ChemStation 环境对 macro 调用字符串可能略有差异。
    这里给几个常见候选写法，方便你切换测试。

    mode='style1'  -> MACRO "C:\\...\\start_single_run.mac"
    mode='style2'  -> MACRO "C:\\...\\start_single_run.mac",GO
    mode='style3'  -> macro "C:\\...\\start_single_run.mac"

    建议：
        先从 style2 开始试，
        如果不行，再改 style1 / style3。
    """
    if mode == "style1":
        return f'MACRO "{macro_path}"'
    elif mode == "style2":
        return f'MACRO "{macro_path}",GO'
    elif mode == "style3":
        return f'macro "{macro_path}"'
    else:
        raise ValueError(f"Unknown macro command style: {mode}")


def start_single_run(
    app_name: str = CHEMSTATION_APP,
    macro_path: str = MACRO_PATH,
    macro_style: str = "style2",
    wait_ready_first: bool = True,
) -> None:
    """
    外部触发一次单次 LC run。

    步骤：
        1. 连接 ChemStation
        2. 可选：等待 READY
        3. 发送 macro 执行命令
        4. 等待状态离开 READY，作为 run 启动的简化确认
    """
    logger.info("Starting single run via ChemStation DDE ...")
    logger.info("App name = %s", app_name)
    logger.info("Macro path = %s", macro_path)

    cs = ChemStationDDE(app_name=app_name)

    if wait_ready_first:
        wait_until_ready(cs)

    command = build_macro_command(macro_path, mode=macro_style)
    cs.execute(TOPIC_EXECUTE, command)

    # 这里不是严格证明“采集已经开始”，
    # 但通常能作为一个很好用的第一层确认。
    wait_until_not_ready(cs)

    logger.info("Single run trigger finished.")


if __name__ == "__main__":
    try:
        start_single_run(
            app_name=CHEMSTATION_APP,
            macro_path=MACRO_PATH,
            macro_style="style2",   # 先试这个
            wait_ready_first=True
        )
    except Exception as e:
        logger.exception("Failed to start single run: %s", e)
        raise
