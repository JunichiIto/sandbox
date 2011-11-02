using NUnit.Framework;

[TestFixture]
public class ExcelUtilTest
{
    private static readonly ExcelUtil excelUtil = new ExcelUtil();

    [Test]
    public void TestToColNumber()
    {
        Assert.AreEqual(1, excelUtil.ToColNumber("A"));
        Assert.AreEqual(2, excelUtil.ToColNumber("B"));
        Assert.AreEqual(3, excelUtil.ToColNumber("C"));
        Assert.AreEqual(25, excelUtil.ToColNumber("Y"));
        Assert.AreEqual(26, excelUtil.ToColNumber("Z"));
        Assert.AreEqual(27, excelUtil.ToColNumber("AA"));
        Assert.AreEqual(28, excelUtil.ToColNumber("AB"));
        Assert.AreEqual(51, excelUtil.ToColNumber("AY"));
        Assert.AreEqual(52, excelUtil.ToColNumber("AZ"));
        Assert.AreEqual(53, excelUtil.ToColNumber("BA"));
        Assert.AreEqual(54, excelUtil.ToColNumber("BB"));
        Assert.AreEqual(256, excelUtil.ToColNumber("IV"));
        Assert.AreEqual(701, excelUtil.ToColNumber("ZY"));
        Assert.AreEqual(702, excelUtil.ToColNumber("ZZ"));
        Assert.AreEqual(703, excelUtil.ToColNumber("AAA"));
        Assert.AreEqual(704, excelUtil.ToColNumber("AAB"));
        Assert.AreEqual(16384, excelUtil.ToColNumber("XFD"));
    }

    [Test]
    public void TestToColString()
    {
        Assert.AreEqual("A", excelUtil.ToColString(1));
        Assert.AreEqual("B", excelUtil.ToColString(2));
        Assert.AreEqual("C", excelUtil.ToColString(3));
        Assert.AreEqual("Y", excelUtil.ToColString(25));
        Assert.AreEqual("Z", excelUtil.ToColString(26));
        Assert.AreEqual("AA", excelUtil.ToColString(27));
        Assert.AreEqual("AB", excelUtil.ToColString(28));
        Assert.AreEqual("AY", excelUtil.ToColString(51));
        Assert.AreEqual("AZ", excelUtil.ToColString(52));
        Assert.AreEqual("BA", excelUtil.ToColString(53));
        Assert.AreEqual("BB", excelUtil.ToColString(54));
        Assert.AreEqual("IV", excelUtil.ToColString(256));
        Assert.AreEqual("ZY", excelUtil.ToColString(701));
        Assert.AreEqual("ZZ", excelUtil.ToColString(702));
        Assert.AreEqual("AAA", excelUtil.ToColString(703));
        Assert.AreEqual("AAB", excelUtil.ToColString(704));
        Assert.AreEqual("XFD", excelUtil.ToColString(16384));
    }
}
