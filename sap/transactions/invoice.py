"""发票事务。"""

from sap.models import SapResult
from sap.session import SapSession


class InvoiceTransaction:
    """发票事务，封装 VF01/VF03 的页面操作。"""

    def __init__(self, session: SapSession):
        """注入共享 SAP 会话。"""
        self.session = session

    def create_proforma(self) -> SapResult:
        """进入 VF01 并创建形式发票。"""
        result = SapResult(step="vf01")
        try:
            self.session.set_text("wnd[0]/tbar[0]/okcd", "/nvf01")
            self.session.send_vkey(0)
            self.session.press("wnd[0]/tbar[0]/btn[11]")
        except Exception as exc:
            return SapResult.fail(f"形式发票添加失败，{exc}", step="vf01")
        return result

    def display_proforma(self) -> SapResult:
        """进入 VF03 并输出当前形式发票号。"""
        result = SapResult(step="vf03")
        try:
            self.session.set_text("wnd[0]/tbar[0]/okcd", "/nvf03")
            self.session.send_vkey(0)
            result.proforma_no = self.session.read_text("wnd[0]/usr/ctxtVBRK-VBELN")
            self.session.find("wnd[0]/mbar/menu[0]/menu[11]").select()
            self.session.press("wnd[1]/tbar[0]/btn[37]")
        except Exception as exc:
            return SapResult.fail(f"形式发票查看失败，{exc}", step="vf03")
        return result
