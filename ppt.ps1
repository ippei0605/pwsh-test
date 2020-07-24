Add-Type -AssemblyName Microsoft.Office.InterOp.PowerPoint
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$app.Presentations.open("/Users/ippei/IdeaProjects/pwsh-test/docs/自己紹介_鈴木_20200715.pptx")
