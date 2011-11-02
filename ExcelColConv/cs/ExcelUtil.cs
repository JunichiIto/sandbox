using System;
using System.Collections.Generic;

public class ExcelUtil
{
    private static readonly int AzLength = (int)'Z' - (int)'A' + 1; // 26
    private static readonly int OffsetNum = (int)'A' - 1; // 64

    #region ToColNumber
    public int ToColNumber(string colStr)
    {
        return ConvertString(0, new Queue<char>(colStr.ToCharArray()));
    }

    private int ConvertString(int ret, Queue<char> chars)
    {
        if (chars.Count == 0)
        {
            return ret;
        }
        else
        {
            char c = chars.Dequeue();
            return ConvertString(CalcDecimal(c, chars.Count) + ret, chars);
        }
    }

    private int CalcDecimal(char c, int times)
    {
        return (int)Math.Pow(AzLength, times) * ((int)c - OffsetNum);
    }
    #endregion

    #region ToColString
    public string ToColString(int colNum)
    {
        return ConvertNumber(string.Empty, colNum);
    }

    private string ToChar(int n)
    {
        return ((char)(n + OffsetNum)).ToString();
    }

    private struct Tuple
    {
        public int quo;
        public int rem;
        public Tuple(int quo, int rem)
        {
            this.quo = quo;
            this.rem = rem;
        }
    }

    private Tuple ExcelDivmod(int n)
    {
        int quo = n / AzLength;
        int rem = n % AzLength;
        if (rem == 0)
        {
            return new Tuple(quo - 1, AzLength);
        }
        else
        {
            return new Tuple(quo, rem);
        }
    }

    private string ConvertNumber(string ret, int n)
    {
        if (n == 0)
        {
            return ret;
        }
        else
        {
            Tuple t = ExcelDivmod(n);
            return ConvertNumber(ToChar(t.rem) + ret, t.quo);
        }
    }
    #endregion
}