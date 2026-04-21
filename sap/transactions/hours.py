"""工时事务。"""

from sap.models import HourData, SapResult
from sap.session import SapSession


class HourTransaction:
    """工时事务，负责登录、录入和保存。"""

    def __init__(self, session: SapSession):
        """注入共享 SAP 会话。"""
        self.session = session

    def login(self, hour: HourData) -> SapResult:
        """登录工时系统并校验员工是否存在。"""
        result = SapResult(step="hour_login")
        try:
            self.session.set_text("wnd[0]/tbar[0]/okcd", "/NZRU1")
            self.session.send_vkey(0)
            self.session.set_text("wnd[0]/usr/ctxtZRUCKD-PERNR", hour.staff_id)
            self.session.set_text("wnd[0]/usr/txtZRUCKD-KWEEK", hour.week)
            self.session.focus("wnd[0]/usr/txtZRUCKD-KWEEK", 2)
            self.session.send_vkey(0)
            self.session.sleep(1)
            status_text = self.session.read_status()
            if "doesn't exist" in status_text or "does not exist" in status_text or "不存在" in status_text:
                return SapResult.fail(f"登录工时系统失败，员工ID无效: {status_text}", step="hour_login")
        except Exception as exc:
            return SapResult.fail(f"Hour界面失败，{exc}", step="hour_login")
        return result

    def record(self, hour: HourData, row_num: int = 0, max_rows: int = 20) -> SapResult:
        """把工时数据写入表格中的第一条空行。"""
        result = SapResult(step="hour_record")
        try:
            # 逐行探测空位，避免覆盖已有工时记录。
            while row_num < max_rows:
                date_id = f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,{row_num}]"
                if self.session.read_text(date_id) == "":
                    break
                row_num += 1
            else:
                return SapResult.fail(f"录Hour失败，未找到可写入的空行，最大行数: {max_rows}", step="hour_record")

            self.session.set_text(date_id, hour.allocated_day)
            self.session.set_text(
                f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/ctxtZRUCKDS-BEARBAUFNR[3,{row_num}]",
                hour.order_no,
            )
            self.session.set_text(
                f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-UEPOS[4,{row_num}]",
                hour.item,
            )
            activity_id = f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/ctxtZRUCKDS-ZZTAETIGNR[9,{row_num}]"
            self.session.set_text(activity_id, hour.material_code)
            self.session.focus(activity_id)
            self.session.set_text(
                f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-PZEIT[13,{row_num}]",
                hour.allocated_hours,
            )
            office_id = f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-BZEIT[15,{row_num}]"
            self.session.set_text(office_id, hour.office_time)
            self.session.focus(office_id, 1)
        except Exception as exc:
            return SapResult.fail(f"录Hour失败，{exc}", step="hour_record")
        return result

    def save(self, max_retries: int = 14) -> SapResult:
        """保存工时，必要时按状态栏结果重试。"""
        result = SapResult(step="hour_save")
        try:
            self.session.press("wnd[0]/tbar[0]/btn[11]")
            self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
            return result
        except Exception as initial_exc:
            retry_count = 0
            last_retry_error = None
            while retry_count < max_retries:
                retry_count += 1
                try:
                    # 某些 SAP 校验失败后，需要连续回车让状态栏先稳定下来。
                    self.session.send_vkey(0)
                    self.session.send_vkey(0)
                    status_text = self.session.read_status()
                    if "Fixed price item is allready fully invoiced" in status_text:
                        continue
                    if "Data was saved" in status_text:
                        result.message = "录Hour成功"
                        return result
                    self.session.press("wnd[0]/tbar[0]/btn[11]")
                    self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
                    return result
                except Exception as retry_exc:
                    last_retry_error = retry_exc
            return SapResult.fail(
                f"保存失败，已重试{max_retries}次。初始错误: {initial_exc}. 最后一次重试错误: {last_retry_error}",
                step="hour_save",
            )
