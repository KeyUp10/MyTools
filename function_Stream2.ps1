#Stream2(ディレクトリ内のファイル名をStream2.txtに書き出す)

function Stream2{

 Get-ChildItem -Name -Recurse | Out-File Stream2.txt

}
