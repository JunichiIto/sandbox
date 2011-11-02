gmcs -target:library -reference:bin/nunit.framework.dll -out:bin/ExcelColConv.dll ExcelUtil.cs ExcelUtilTest.cs 
mono bin/nunit-console.exe bin/ExcelColConv.dll 
