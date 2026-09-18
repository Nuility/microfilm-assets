Option Explicit
Dim args, inputPath, outputPath, wordApp, doc
Set args = WScript.Arguments
inputPath = args(0)
outputPath = args(1)
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False
wordApp.DisplayAlerts = 0
Set doc = wordApp.Documents.Open(inputPath, False, True, False)
doc.ExportAsFixedFormat outputPath, 17, False, 0, 0, 1, doc.ComputeStatistics(2), 0, True, True, 0, True, True, False
doc.Close False
wordApp.Quit
WScript.Echo outputPath
