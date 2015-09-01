With CreateObject("Excel.Application")
  .Workbooks.Open(WScript.CreateObject("WScript.Shell").CurrentDirectory & "\bot.xlsm")
  .run "bot"
  .Quit
End With

