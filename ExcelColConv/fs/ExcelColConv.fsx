namespace ExcelColConv
module ExcelUtil = 
  let AzLength = int 'Z' - int 'A' + 1 // 26
  let OffsetNum = int 'A' - 1 // 64
  
  let toColNumber (colStr : string) =
    let calcDecimal c times = (pown AzLength times) * (int c - OffsetNum) 
    let rec convertStr ret = function
      | [] -> ret
      | c::chars -> ((calcDecimal c chars.Length) + ret, chars)
                    ||> convertStr
    convertStr 0 (Seq.toList colStr)
  
  let toColString colNum =
    let toChar n = n + OffsetNum |> char |> string
    let excelDivmod n = 
      match n / AzLength, n % AzLength with
      | quo, 0 -> quo - 1, AzLength
      | t -> t
    let rec convertNum ret = function
      | 0 -> ret
      | n -> excelDivmod n 
             ||> fun quo rem -> (toChar rem) + ret, quo
             ||> convertNum
    convertNum "" colNum
