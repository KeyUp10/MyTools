###使用方法###
#
# 1) 基本(単発)
# URL-Codec encode "日本語"
# URL-Codec decode "%E6%97%A5%E6%9C%AC%E8%AA%9E"
#
# 2) パイプで複数行
# "日本語","テスト","PowerShell" | URL-Codec encode
#
# 3) ファイルから読み込んで変換
# Get-Content words.txt | URL-Codec decode
#
###使用方法###

function URL-Codec {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("encode","decode")]
        [string]$Mode,

        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
        [string]$Text
    )

    process {
        switch ($Mode) {
            "encode" { [System.Net.WebUtility]::UrlEncode($Text) }
            "decode" { [System.Net.WebUtility]::UrlDecode($Text) }
        }
    }
}
