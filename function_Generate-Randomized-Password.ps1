#ランダムな文字列を生成する関数

Add-type -AssemblyName System.Web

function Generate-Randomized-Password([int]$a, [int]$b) {

     [System.Web.Security.Membership]::GeneratePassword($a, $b)

}
