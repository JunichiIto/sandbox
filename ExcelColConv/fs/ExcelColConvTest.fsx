#r "Onigiri.dll"
#load @"ExcelColConv.fsx"

open Onigiri
open Onigiri.Assertion.Operators
open ExcelColConv

type MyTest() =
    [<Test>] 
    let test_equal () = 1 == ExcelUtil.toColNumber "A"

    [<Test>] 
    let test_equal () = "A" == ExcelUtil.toColString 1 

let run () = onigiri<MyTest>
do run ()
