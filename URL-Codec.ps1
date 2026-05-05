param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("encode","decode")]
    [string]$Mode,

    [Parameter(Mandatory=$true)]
    [string]$Text
)

switch ($Mode) {
    "encode" { [System.Net.WebUtility]::UrlEncode($Text) }
    "decode" { [System.Net.WebUtility]::UrlDecode($Text) }
}