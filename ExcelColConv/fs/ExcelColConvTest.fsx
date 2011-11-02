#r "Onigiri.dll"
#load @"ExcelColConv.fsx"

open Onigiri
open Onigiri.Assertion.Operators
open ExcelColConv

type MyTest() =
  [<Test>] 
  let test_equal () = 1 == ExcelUtil.toColNumber "A"
  [<Test>] 
  let test_equal () = 2 == ExcelUtil.toColNumber "B"
  [<Test>] 
  let test_equal () = 25 == ExcelUtil.toColNumber "Y"
  [<Test>] 
  let test_equal () = 26 == ExcelUtil.toColNumber "Z"
  [<Test>] 
  let test_equal () = 27 == ExcelUtil.toColNumber "AA"
  [<Test>] 
  let test_equal () = 28 == ExcelUtil.toColNumber "AB"
  [<Test>] 
  let test_equal () = 51 == ExcelUtil.toColNumber "AY"
  [<Test>] 
  let test_equal () = 52 == ExcelUtil.toColNumber "AZ"
  [<Test>] 
  let test_equal () = 53 == ExcelUtil.toColNumber "BA"
  [<Test>] 
  let test_equal () = 54 == ExcelUtil.toColNumber "BB"
  [<Test>] 
  let test_equal () = 701 == ExcelUtil.toColNumber "ZY"
  [<Test>] 
  let test_equal () = 702 == ExcelUtil.toColNumber "ZZ"
  [<Test>] 
  let test_equal () = 703 == ExcelUtil.toColNumber "AAA"
  [<Test>] 
  let test_equal () = 704 == ExcelUtil.toColNumber "AAB"
  [<Test>] 
  let test_equal () = 16384 == ExcelUtil.toColNumber "XFD"
  
  [<Test>] 
  let test_equal () = "A" == ExcelUtil.toColString 1
  [<Test>] 
  let test_equal () = "B" == ExcelUtil.toColString 2
  [<Test>] 
  let test_equal () = "Y" == ExcelUtil.toColString 25
  [<Test>] 
  let test_equal () = "Z" == ExcelUtil.toColString 26
  [<Test>] 
  let test_equal () = "AA" == ExcelUtil.toColString 27
  [<Test>] 
  let test_equal () = "AB" == ExcelUtil.toColString 28
  [<Test>] 
  let test_equal () = "AY" == ExcelUtil.toColString 51
  [<Test>] 
  let test_equal () = "AZ" == ExcelUtil.toColString 52
  [<Test>] 
  let test_equal () = "BA" == ExcelUtil.toColString 53
  [<Test>] 
  let test_equal () = "BB" == ExcelUtil.toColString 54
  [<Test>] 
  let test_equal () = "ZY" == ExcelUtil.toColString 701
  [<Test>] 
  let test_equal () = "ZZ" == ExcelUtil.toColString 702
  [<Test>] 
  let test_equal () = "AAA" == ExcelUtil.toColString 703
  [<Test>] 
  let test_equal () = "AAB" == ExcelUtil.toColString 704
  [<Test>] 
  let test_equal () = "XFD" == ExcelUtil.toColString 16384

let run () = onigiri<MyTest>
do run ()
