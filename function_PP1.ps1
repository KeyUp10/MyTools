#PP1(PP.txtのデータを縦1行に並べる)

function PP1{

 $str = Get-Content PP.txt
 $str = $str.Replace(' ', '')
 $array = $str.split(";")
 $array | Out-File PP1.txt
}
