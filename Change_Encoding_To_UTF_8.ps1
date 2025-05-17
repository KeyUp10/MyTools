#エンコーディングを UTF-8 に変更する
[console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Output "エンコーディングを変更しました。現在の出力エンコーディングは: $([console]::OutputEncoding.EncodingName)"

