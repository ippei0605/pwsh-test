# pwsh-test
これは PowerShell の勉強です。

## やったこと

### Mac に PowerShell をインストールする
```shell script
brew cask install powershell
```
```shell script
pwsh
```
> 起動成功、コマンドだけでなく LaunchPad からも起動できる。

### IntelliJ IDEA に PowerShell プラグインをインストールする。
<img src="./docs/plugin.png" width="500px">

> .ps1 を作成したら、IntelliJ IDEA がプラグインをインストールするようにおすすめしてくれる。

<img src="./docs/warning_alias.png" width="500px">

> エイリアスで書くと警告されるようだ。

<img src="./docs/code_assist.png" width="500px">

> コードアシストしてくれる。

### Hello World
hello.ps1 を実行してみる。
```
% pwsh
PowerShell 7.0.3
Copyright (c) Microsoft Corporation. All rights reserved.

https://aka.ms/powershell
Type 'help' to get help.

PS /Users/ippei/IdeaProjects/pwsh-test>./hello.ps1
Hello World
Hello World
```

### PowerPoint を開く
```
PS /Users/ippei/IdeaProjects/pwsh-test> ./ppt.ps1
Add-Type: /Users/ippei/IdeaProjects/pwsh-test/ppt.ps1:1
Line |
   1 |  Add-Type -AssemblyName Microsoft.Office.InterOp.PowerPoint
     |  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     | Cannot find path '/Users/ippei/IdeaProjects/pwsh-test/Microsoft.Office.InterOp.PowerPoint.dll' because it does not exist.

New-Object: /Users/ippei/IdeaProjects/pwsh-test/ppt.ps1:2
Line |
   2 |  $app = New-Object -ComObject PowerPoint.Application
     |                    ~~~~~~~~~~
     | A parameter cannot be found that matches parameter name 'ComObject'.

InvalidOperation: /Users/ippei/IdeaProjects/pwsh-test/ppt.ps1:3
Line |
   3 |  $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
     |                 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     | Unable to find type [Microsoft.Office.Core.MsoTriState].

InvalidOperation: /Users/ippei/IdeaProjects/pwsh-test/ppt.ps1:4
Line |
   4 |  $app.Presentations.open("/Users/ippei/IdeaProjects/pwsh-test/docs/自己紹 …
     |  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
     | You cannot call a method on a null-valued expression.
```

> Mac だと Microsoft Office を操作できない模様。

## まとめ？
快適な Mac の開発環境で PowerShell を勉強できると思ったが、あっけなく終了。

## 参考情報
https://docs.microsoft.com/ja-jp/powershell/scripting/install/installing-powershell-core-on-macos?view=powershell-7
