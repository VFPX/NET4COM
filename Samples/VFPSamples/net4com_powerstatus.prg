CLEAR
LOCAL x as NET4COM.IPowerStatus
x=CREATEOBJECT("NET4COM.PowerStatus")
? "BatteryLifePercent: "+STR(x.BatteryLifePercent)
? "BatteryFullLifetime: "+STR(x.BatteryFullLifetime)
? "BatteryLifeRemaining: "+STR(x.BatteryLifeRemaining)
? "BatteryChargeStatus: "+STR(x.BatteryChargeStatus)
? "PowerLineStatus: "+STR(x.PowerLineStatus)
RETURN
