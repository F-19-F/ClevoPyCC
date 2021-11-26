import mi
import time


class WMI_CLS():
    def __init__(self, namespace, cls_name) -> None:
        self.app = mi.Application()
        self.session = self.app.create_session(protocol=mi.PROTOCOL_WMIDCOM)
        try:
            self.WMI_QUERY = self.session.exec_query(namespace, cls_name)
            self.WMI_INS = self.WMI_QUERY.get_next_instance()
        except Exception as e:
            # MI_RESULT_ACCESS_DENIED
            if (e.args[0]['mi_result'] == 2):
                print("run as administrator!")
                exit(-1)
            raise e
        self.last_opt = None
        self.cls = self.WMI_INS.get_class()

    def EvaluateMethod(self, method_name, **kwargs):
        if self.last_opt:
            self.last_opt.close()
        params = self.app.create_method_params(self.cls, method_name)
        if len(kwargs) != 0:
            for k in kwargs:
                params[k] = kwargs[k]
        _call = self.session.invoke_method(self.WMI_INS, method_name, params)
        self.last_opt = _call
        if _call.has_more_results():
            res = _call.get_next_instance()
            return res.clone()

    # stop threads
    def __del__(self):
        if 'last_opt' in self.__dict__ and self.last_opt:
            self.last_opt.close()
        if 'WMI_QUERY' in self.__dict__:
            self.WMI_QUERY.close()
        if 'session' in self.__dict__:
            self.session.close()
        self.app.close()


class CLEVO_GET():
    def __init__(self) -> None:
        self.cls = WMI_CLS("root\WMI", "select * from CLEVO_GET")

    def GetFanInfo(self, fanNr: int):
        if fanNr < 1 or fanNr > 4:
            return None
        res = self.cls.EvaluateMethod("Fan{}Info".format(fanNr))
        res = res['Data']
        result = {
            "raw_percent": res & 0xFF,
            "rpm_percent": int(((res & 0xFF)/0xFF)*100+0.5),
            "temperature": res >> 0x10
        }
        return result

    def GetFanRpm(self, fanNr: int):
        maps = {
            1: '12',
            2: '12',
            3: '34',
            4: '34'
        }
        if fanNr < 1 or fanNr > 4:
            return None
        res = self.cls.EvaluateMethod("GetFan{}RPM".format(maps[fanNr]))
        # 32bits
        res = res['Data']
        if fanNr > int(maps[fanNr][0]):
            # high 16 bits in LE
            if res >> 0x10 == 0:
                return 0
            return int(2156220/(res >> 0x10)+0.5)
        else:
            if res & 0x0000FFFF == 0:
                return 0
            # low 16 bits in LE
            return int(2156220/(res & 0x0000FFFF)+0.5)

    def SetFanSpeedPercent(self, fanNr: int, percent: int):
        if percent > 100 or percent < 0:
            return None
        args = 0
        for i in range(4):
            if i+1 == fanNr:
                # target
                args |= (int((percent/100)*0xFF+0.5) << 8*i)
            else:
                args |= (self.GetFanInfo(i+1)['raw_percent'] << 8*i)
        self.cls.EvaluateMethod("SetFanDuty", Data=args)

    def SetFansAuto(self):
        args = 0
        args |= 1
        args |= 1 << 0x01
        args |= 1 << 0x02
        args |= 1 << 0x03
        self.cls.EvaluateMethod("SetFanAutoDuty", Data=args)

    def SetWhiteLedKB(self, level: int):
        self.cls.EvaluateMethod("SetWhiteLedKB", Data=level)

    def GetWhiteLedKB(self):
        res = self.cls.EvaluateMethod("GetWhiteLedKB")
        return res['Data']


if __name__ == '__main__':
    CG = CLEVO_GET()
    print("---FAN INFO---")
    info1 = CG.GetFanInfo(1)
    info1.update({"rpm": CG.GetFanRpm(1)})
    info2 = CG.GetFanInfo(2)
    info2.update({"rpm": CG.GetFanRpm(2)})
    print("NO.1 rpm:{rpm}\tpercent:{rpm_percent}%\ttemperature:{temperature}".format(
        **info1))
    print("NO.2 rpm:{rpm}\tpercent:{rpm_percent}%\ttemperature:{temperature}".format(
        **info2))
    print("set fan 1 speed to max")
    CG.SetFanSpeedPercent(1, 100)
    time.sleep(1)
    print("set fan 2 speed to max")
    CG.SetFanSpeedPercent(2, 100)
    time.sleep(1)
    print("set fan speed to auto")
    CG.SetFansAuto()
    for i in range(6):
        print("set WhiteLedKB to {}".format(i))
        CG.SetWhiteLedKB(i)
        time.sleep(1)
