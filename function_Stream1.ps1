#Stream1(ディレクトリ内のファイル名をStream1.txtに書き出す)

function Stream1{

 Get-ChildItem -Recurse | Out-File Stream1.txt
}
