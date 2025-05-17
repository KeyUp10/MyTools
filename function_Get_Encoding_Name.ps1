function Get-Encoding-Name{

#現在の出力エンコーディングの名前を取得する
$encodingName = [console]::OutputEncoding.EncodingName

#結果を表示する
Write-Output "現在の出力エンコーディングは: $encodingName"

}
