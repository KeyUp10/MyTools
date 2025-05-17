function Generate-Randomized-Password([int]$length, [int]$symbolCount) {

    $lettersAndNumbers = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    $symbols = "!@#$%^&*()-=_+"
    
    #ランダムに文字と数字を選択
    $passwordCore = -join ((1..($length - $symbolCount)) | ForEach-Object { $lettersAndNumbers[(Get-Random -Maximum $lettersAndNumbers.Length)] })
    
    #ランダムに記号を選択
    $symbolPart = -join ((1..$symbolCount) | ForEach-Object { $symbols[(Get-Random -Maximum $symbols.Length)] })
    
    #パスワードをシャッフルして返す
    $password = ($passwordCore + $symbolPart).ToCharArray() | Sort-Object {Get-Random} -Unique

    return -join $password

}
