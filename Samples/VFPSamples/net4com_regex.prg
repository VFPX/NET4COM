CLEAR
LOCAL x as NET4COM.IRegEx
x=CREATEOBJECT("NET4COM.RegEx")
? "String to convert:  4255551234"
? x.Replace("4255551234","(\d{3})(\d{3})(\d{4})","($1) $2-$3")
? "-------"
LOCAL cEmailAddress
cEmailAddress="foobar@msdn.com"
? cEmailAddress+IIF(x.IsMatch(cEmailAddress,"^(?<user>[^@]+)@(?<host>.+)$")," is a valid email"," is NOT a valid email")
cEmailAddress="foobarmsdn.com"
? cEmailAddress+IIF(x.IsMatch(cEmailAddress,"^(?<user>[^@]+)@(?<host>.+)$"), ;
		" is a valid email"," is NOT a valid email")
? "-------"
? "Match string: "+"One car red car blue car"
LOCAL a1,a2
a1=x.Matches("One car red car blue car","(\w+)\s+(car)")
FOR num = 1 TO ALEN(a1)
	? a1(num)
ENDFOR
? "-------"
? "Split string: "+"one-two-three"
a2=x.Split("one-two-three","(-)")
FOR num = 1 TO ALEN(a2)
	? a2(num)
ENDFOR
RETURN
