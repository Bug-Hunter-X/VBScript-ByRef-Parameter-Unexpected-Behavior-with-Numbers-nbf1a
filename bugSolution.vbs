Function f(ByVal a)
  a = a + 1
  f = a 'Return the new value
end function

x = 10
y = f(x)
MsgBox x & ", " & y