using System;
using System.Text.RegularExpressions;

namespace XLibur.Excel.CalcEngine.Functions;

internal sealed class XLMatrix
{
    public XLMatrix? L;
    public XLMatrix? U;
    public int Cols;
    private double _detOfP = 1;
    public double[,] Mat;
    private int[]? _pi;
    private readonly int _rows;

    public XLMatrix(int iRows, int iCols) // XLMatrix Class constructor
    {
        _rows = iRows;
        Cols = iCols;
        Mat = new double[_rows, Cols];
    }

    public XLMatrix(double[,] arr)
        : this(arr.GetLength(0), arr.GetLength(1))
    {
        var roCount = arr.GetLength(0);
        var coCount = arr.GetLength(1);
        for (int ro = 0; ro < roCount; ro++)
        {
            for (int co = 0; co < coCount; co++)
            {
                Mat[ro, co] = arr[ro, co];
            }
        }
    }

    public double this[int iRow, int iCol] // Access this matrix as a 2D array
    {
        get { return Mat[iRow, iCol]; }
        set { Mat[iRow, iCol] = value; }
    }

    public bool IsSingular()
    {
        for (var row = 0; row < _rows; row++)
        {
            for (var col = 0; col < Cols; col++)
            {
                var element = Mat[row, col];
                if (double.IsNaN(element) || double.IsInfinity(element))
                    return true;
            }
        }

        return false;
    }

    public bool IsSquare()
    {
        return (_rows == Cols);
    }

    public XLMatrix GetCol(int k)
    {
        var m = new XLMatrix(_rows, 1);
        for (var i = 0; i < _rows; i++) m[i, 0] = Mat[i, k];
        return m;
    }

    public void SetCol(XLMatrix v, int k)
    {
        for (var i = 0; i < _rows; i++) Mat[i, k] = v[i, 0];
    }

    public void MakeLu() // Function for LU decomposition
    {
        if (!IsSquare()) throw new InvalidOperationException("The matrix is not square!");
        L = IdentityMatrix(_rows, Cols);
        U = Duplicate();

        _pi = new int[_rows];
        for (var i = 0; i < _rows; i++) _pi[i] = i;

        var k0 = 0;

        for (var k = 0; k < Cols - 1; k++)
        {
            k0 = FindPivotRow(k);

            var pom1 = _pi[k];
            _pi[k] = _pi[k0];
            _pi[k0] = pom1; // switch two rows in permutation matrix

            SwapLuRows(k, k0);

            for (var i = k + 1; i < _rows; i++)
            {
                L[i, k] = U[i, k] / U[k, k];
                for (var j = k; j < Cols; j++)
                    U[i, j] = U[i, j] - L[i, k] * U[k, j];
            }
        }
    }

    private int FindPivotRow(int k)
    {
        double p = 0;
        var k0 = k;
        for (var i = k; i < _rows; i++)
        {
            if (Math.Abs(U![i, k]) > p)
            {
                p = Math.Abs(U[i, k]);
                k0 = i;
            }
        }
        if (p == 0)
            throw new InvalidOperationException("The matrix is singular!");
        return k0;
    }

    private void SwapLuRows(int k, int k0)
    {
        double pom2;
        for (var i = 0; i < k; i++)
        {
            pom2 = L![k, i];
            L[k, i] = L[k0, i];
            L[k0, i] = pom2;
        }

        if (k != k0) _detOfP *= -1;

        for (var i = 0; i < Cols; i++)
        {
            pom2 = U![k, i];
            U[k, i] = U[k0, i];
            U[k0, i] = pom2;
        }
    }


    public XLMatrix SolveWith(XLMatrix v) // Function solves Ax = v in conformity with solution vector "v"
    {
        if (_rows != Cols) throw new InvalidOperationException("The matrix is not square!");
        if (_rows != v._rows) throw new ArgumentException("Wrong number of results in solution vector!");
        if (L == null) MakeLu();

        var b = new XLMatrix(_rows, 1);
        for (var i = 0; i < _rows; i++) b[i, 0] = v[_pi![i], 0]; // switch two items in "v" due to permutation matrix

        var z = SubsForth(L!, b);
        var x = SubsBack(U!, z);

        return x;
    }

    public XLMatrix Invert() // Function returns the inverted matrix
    {
        if (L == null) MakeLu();

        var inv = new XLMatrix(_rows, Cols);

        for (var i = 0; i < _rows; i++)
        {
            var ei = ZeroMatrix(_rows, 1);
            ei[i, 0] = 1;
            var col = SolveWith(ei);
            inv.SetCol(col, i);
        }
        return inv;
    }

    public double Determinant() // Function for determinant
    {
        if (L == null) MakeLu();
        var det = _detOfP;
        for (var i = 0; i < _rows; i++) det *= U![i, i];
        return det;
    }

    public XLMatrix GetP() // Function returns permutation matrix "P" due to permutation vector "pi"
    {
        if (L == null) MakeLu();

        var matrix = ZeroMatrix(_rows, Cols);
        for (var i = 0; i < _rows; i++) matrix[_pi![i], i] = 1;
        return matrix;
    }

    public XLMatrix Duplicate() // Function returns the copy of this matrix
    {
        var matrix = new XLMatrix(_rows, Cols);
        for (var i = 0; i < _rows; i++)
            for (var j = 0; j < Cols; j++)
                matrix[i, j] = Mat[i, j];
        return matrix;
    }

    public static XLMatrix SubsForth(XLMatrix a, XLMatrix b) // Function solves Ax = b for A as a lower triangular matrix
    {
        if (a.L == null) a.MakeLu();
        var n = a._rows;
        var x = new XLMatrix(n, 1);

        for (var i = 0; i < n; i++)
        {
            x[i, 0] = b[i, 0];
            for (var j = 0; j < i; j++) x[i, 0] -= a[i, j] * x[j, 0];
            x[i, 0] = x[i, 0] / a[i, i];
        }
        return x;
    }

    public static XLMatrix SubsBack(XLMatrix a, XLMatrix b) // Function solves Ax = b for A as an upper triangular matrix
    {
        if (a.L == null) a.MakeLu();
        var n = a._rows;
        var x = new XLMatrix(n, 1);

        for (var i = n - 1; i > -1; i--)
        {
            x[i, 0] = b[i, 0];
            for (var j = n - 1; j > i; j--) x[i, 0] -= a[i, j] * x[j, 0];
            x[i, 0] = x[i, 0] / a[i, i];
        }
        return x;
    }

    public static XLMatrix ZeroMatrix(int iRows, int iCols) // Function generates the zero matrix
    {
        var matrix = new XLMatrix(iRows, iCols);
        for (var i = 0; i < iRows; i++)
            for (var j = 0; j < iCols; j++)
                matrix[i, j] = 0;
        return matrix;
    }

    public static XLMatrix IdentityMatrix(int iRows, int iCols) // Function generates the identity matrix
    {
        var matrix = ZeroMatrix(iRows, iCols);
        for (var i = 0; i < Math.Min(iRows, iCols); i++)
            matrix[i, i] = 1;
        return matrix;
    }

    public static XLMatrix RandomMatrix(int iRows, int iCols, int dispersion) // Function generates the zero matrix
    {
        var random = new Random();
        var matrix = new XLMatrix(iRows, iCols);
        for (var i = 0; i < iRows; i++)
            for (var j = 0; j < iCols; j++)
                matrix[i, j] = random.Next(-dispersion, dispersion);
        return matrix;
    }

    public static XLMatrix Parse(string ps) // Function parses the matrix from string
    {
        var s = NormalizeMatrixString(ps);
        var rows = Regex.Split(s, "\r\n");
        var nums = rows[0].Split(' ');
        var matrix = new XLMatrix(rows.Length, nums.Length);
        try
        {
            for (var i = 0; i < rows.Length; i++)
            {
                nums = rows[i].Split(' ');
                for (var j = 0; j < nums.Length; j++) matrix[i, j] = double.Parse(nums[j]);
            }
        }
        catch (FormatException fe)
        {
            throw new FormatException("Wrong input format!", fe);
        }
        return matrix;
    }

    public override string ToString() // Function returns matrix as a string
    {
        var s = "";
        for (var i = 0; i < _rows; i++)
        {
            for (var j = 0; j < Cols; j++) s += $"{Mat[i, j],5:0.00}" + " ";
            s += "\r\n";
        }
        return s;
    }

    public static XLMatrix Transpose(XLMatrix m) // XLMatrix transpose, for any rectangular matrix
    {
        var t = new XLMatrix(m.Cols, m._rows);
        for (var i = 0; i < m._rows; i++)
            for (var j = 0; j < m.Cols; j++)
                t[j, i] = m[i, j];
        return t;
    }

    public static XLMatrix Power(XLMatrix m, int pow) // Power matrix to exponent
    {
        if (pow == 0) return IdentityMatrix(m._rows, m.Cols);
        if (pow == 1) return m.Duplicate();
        if (pow == -1) return m.Invert();

        XLMatrix x;
        if (pow < 0)
        {
            x = m.Invert();
            pow *= -1;
        }
        else x = m.Duplicate();

        var ret = IdentityMatrix(m._rows, m.Cols);
        while (pow != 0)
        {
            if ((pow & 1) == 1) ret *= x;
            x *= x;
            pow >>= 1;
        }
        return ret;
    }

    private static void SafeAplusBintoC(XLMatrix a, int xa, int ya, XLMatrix b, int xb, int yb, XLMatrix c, int size)
    {
        for (var i = 0; i < size; i++) // rows
            for (var j = 0; j < size; j++) // cols
            {
                c[i, j] = 0;
                if (xa + j < a.Cols && ya + i < a._rows) c[i, j] += a[ya + i, xa + j];
                if (xb + j < b.Cols && yb + i < b._rows) c[i, j] += b[yb + i, xb + j];
            }
    }

    private static void SafeAminusBintoC(XLMatrix a, int xa, int ya, XLMatrix b, int xb, int yb, XLMatrix c, int size)
    {
        for (var i = 0; i < size; i++) // rows
            for (var j = 0; j < size; j++) // cols
            {
                c[i, j] = 0;
                if (xa + j < a.Cols && ya + i < a._rows) c[i, j] += a[ya + i, xa + j];
                if (xb + j < b.Cols && yb + i < b._rows) c[i, j] -= b[yb + i, xb + j];
            }
    }

    private static void SafeACopytoC(XLMatrix a, int xa, int ya, XLMatrix c, int size)
    {
        for (var i = 0; i < size; i++) // rows
            for (var j = 0; j < size; j++) // cols
            {
                c[i, j] = 0;
                if (xa + j < a.Cols && ya + i < a._rows) c[i, j] += a[ya + i, xa + j];
            }
    }

    private static void AplusBintoC(XLMatrix a, int xa, int ya, XLMatrix b, int xb, int yb, XLMatrix c, int size)
    {
        for (var i = 0; i < size; i++) // rows
            for (var j = 0; j < size; j++) c[i, j] = a[ya + i, xa + j] + b[yb + i, xb + j];
    }

    private static void AminusBintoC(XLMatrix a, int xa, int ya, XLMatrix b, int xb, int yb, XLMatrix c, int size)
    {
        for (var i = 0; i < size; i++) // rows
            for (var j = 0; j < size; j++) c[i, j] = a[ya + i, xa + j] - b[yb + i, xb + j];
    }

    private static void ACopytoC(XLMatrix a, int xa, int ya, XLMatrix c, int size)
    {
        for (var i = 0; i < size; i++) // rows
            for (var j = 0; j < size; j++) c[i, j] = a[ya + i, xa + j];
    }

    private static XLMatrix StrassenMultiply(XLMatrix a, XLMatrix b) // Smart matrix multiplication
    {
        if (a.Cols != b._rows) throw new ArgumentException("Wrong dimension of matrix!");

        var msize = Math.Max(Math.Max(a._rows, a.Cols), Math.Max(b._rows, b.Cols));

        if (msize < 32)
            return NaiveMultiply(a, b);

        var size = 1;
        var n = 0;
        while (msize > size)
        {
            size *= 2;
            n++;
        }

        var h = size / 2;

        var mField = new XLMatrix[n, 9];

        for (var i = 0; i < n - 4; i++)
        {
            var z = (int)Math.Pow(2, n - i - 1);
            for (var j = 0; j < 9; j++) mField[i, j] = new XLMatrix(z, z);
        }

        StrassenComputeProducts(a, b, h, mField);

        var r = new XLMatrix(a._rows, b.Cols);
        StrassenAssembleResult(r, h, mField);
        return r;
    }

    private static XLMatrix NaiveMultiply(XLMatrix a, XLMatrix b)
    {
        var r = ZeroMatrix(a._rows, b.Cols);
        for (var i = 0; i < r._rows; i++)
            for (var j = 0; j < r.Cols; j++)
                for (var k = 0; k < a.Cols; k++)
                    r[i, j] += a[i, k] * b[k, j];
        return r;
    }

    private static void StrassenComputeProducts(XLMatrix a, XLMatrix b, int h, XLMatrix[,] mField)
    {
        SafeAplusBintoC(a, 0, 0, a, h, h, mField[0, 0], h);
        SafeAplusBintoC(b, 0, 0, b, h, h, mField[0, 1], h);
        StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 1], 1, mField);

        SafeAplusBintoC(a, 0, h, a, h, h, mField[0, 0], h);
        SafeACopytoC(b, 0, 0, mField[0, 1], h);
        StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 2], 1, mField);

        SafeACopytoC(a, 0, 0, mField[0, 0], h);
        SafeAminusBintoC(b, h, 0, b, h, h, mField[0, 1], h);
        StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 3], 1, mField);

        SafeACopytoC(a, h, h, mField[0, 0], h);
        SafeAminusBintoC(b, 0, h, b, 0, 0, mField[0, 1], h);
        StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 4], 1, mField);

        SafeAplusBintoC(a, 0, 0, a, h, 0, mField[0, 0], h);
        SafeACopytoC(b, h, h, mField[0, 1], h);
        StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 5], 1, mField);

        SafeAminusBintoC(a, 0, h, a, 0, 0, mField[0, 0], h);
        SafeAplusBintoC(b, 0, 0, b, h, 0, mField[0, 1], h);
        StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 6], 1, mField);

        SafeAminusBintoC(a, h, 0, a, h, h, mField[0, 0], h);
        SafeAplusBintoC(b, 0, h, b, h, h, mField[0, 1], h);
        StrassenMultiplyRun(mField[0, 0], mField[0, 1], mField[0, 1 + 7], 1, mField);
    }

    private static void StrassenAssembleResult(XLMatrix r, int h, XLMatrix[,] mField)
    {
        for (var i = 0; i < Math.Min(h, r._rows); i++)
            for (var j = 0; j < Math.Min(h, r.Cols); j++)
                r[i, j] = mField[0, 1 + 1][i, j] + mField[0, 1 + 4][i, j] - mField[0, 1 + 5][i, j] +
                          mField[0, 1 + 7][i, j];

        for (var i = 0; i < Math.Min(h, r._rows); i++)
            for (var j = h; j < Math.Min(2 * h, r.Cols); j++)
                r[i, j] = mField[0, 1 + 3][i, j - h] + mField[0, 1 + 5][i, j - h];

        for (var i = h; i < Math.Min(2 * h, r._rows); i++)
            for (var j = 0; j < Math.Min(h, r.Cols); j++)
                r[i, j] = mField[0, 1 + 2][i - h, j] + mField[0, 1 + 4][i - h, j];

        for (var i = h; i < Math.Min(2 * h, r._rows); i++)
            for (var j = h; j < Math.Min(2 * h, r.Cols); j++)
                r[i, j] = mField[0, 1 + 1][i - h, j - h] - mField[0, 1 + 2][i - h, j - h] +
                          mField[0, 1 + 3][i - h, j - h] + mField[0, 1 + 6][i - h, j - h];
    }

    // function for square matrix 2^N x 2^N

    private static void StrassenMultiplyRun(XLMatrix a, XLMatrix b, XLMatrix c, int l, XLMatrix[,] f)
    {
        var size = a._rows;
        var h = size / 2;

        if (size < 32)
        {
            NaiveMultiplyInto(a, b, c);
            return;
        }

        StrassenRunComputeProducts(a, b, h, l, f);
        StrassenRunAssembleResult(c, h, size, l, f);
    }

    private static void NaiveMultiplyInto(XLMatrix a, XLMatrix b, XLMatrix c)
    {
        for (var i = 0; i < c._rows; i++)
            for (var j = 0; j < c.Cols; j++)
            {
                c[i, j] = 0;
                for (var k = 0; k < a.Cols; k++) c[i, j] += a[i, k] * b[k, j];
            }
    }

    private static void StrassenRunComputeProducts(XLMatrix a, XLMatrix b, int h, int l, XLMatrix[,] f)
    {
        AplusBintoC(a, 0, 0, a, h, h, f[l, 0], h);
        AplusBintoC(b, 0, 0, b, h, h, f[l, 1], h);
        StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 1], l + 1, f);

        AplusBintoC(a, 0, h, a, h, h, f[l, 0], h);
        ACopytoC(b, 0, 0, f[l, 1], h);
        StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 2], l + 1, f);

        ACopytoC(a, 0, 0, f[l, 0], h);
        AminusBintoC(b, h, 0, b, h, h, f[l, 1], h);
        StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 3], l + 1, f);

        ACopytoC(a, h, h, f[l, 0], h);
        AminusBintoC(b, 0, h, b, 0, 0, f[l, 1], h);
        StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 4], l + 1, f);

        AplusBintoC(a, 0, 0, a, h, 0, f[l, 0], h);
        ACopytoC(b, h, h, f[l, 1], h);
        StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 5], l + 1, f);

        AminusBintoC(a, 0, h, a, 0, 0, f[l, 0], h);
        AplusBintoC(b, 0, 0, b, h, 0, f[l, 1], h);
        StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 6], l + 1, f);

        AminusBintoC(a, h, 0, a, h, h, f[l, 0], h);
        AplusBintoC(b, 0, h, b, h, h, f[l, 1], h);
        StrassenMultiplyRun(f[l, 0], f[l, 1], f[l, 1 + 7], l + 1, f);
    }

    private static void StrassenRunAssembleResult(XLMatrix c, int h, int size, int l, XLMatrix[,] f)
    {
        for (var i = 0; i < h; i++)
            for (var j = 0; j < h; j++)
                c[i, j] = f[l, 1 + 1][i, j] + f[l, 1 + 4][i, j] - f[l, 1 + 5][i, j] + f[l, 1 + 7][i, j];

        for (var i = 0; i < h; i++)
            for (var j = h; j < size; j++)
                c[i, j] = f[l, 1 + 3][i, j - h] + f[l, 1 + 5][i, j - h];

        for (var i = h; i < size; i++)
            for (var j = 0; j < h; j++)
                c[i, j] = f[l, 1 + 2][i - h, j] + f[l, 1 + 4][i - h, j];

        for (var i = h; i < size; i++)
            for (var j = h; j < size; j++)
                c[i, j] = f[l, 1 + 1][i - h, j - h] - f[l, 1 + 2][i - h, j - h] + f[l, 1 + 3][i - h, j - h] +
                          f[l, 1 + 6][i - h, j - h];
    }

    public static XLMatrix StupidMultiply(XLMatrix m1, XLMatrix m2) // Stupid matrix multiplication
    {
        if (m1.Cols != m2._rows) throw new ArgumentException("Wrong dimensions of matrix!");

        var result = ZeroMatrix(m1._rows, m2.Cols);
        for (var i = 0; i < result._rows; i++)
            for (var j = 0; j < result.Cols; j++)
                for (var k = 0; k < m1.Cols; k++)
                    result[i, j] += m1[i, k] * m2[k, j];
        return result;
    }

    private static XLMatrix Multiply(double n, XLMatrix m) // Multiplication by constant n
    {
        var r = new XLMatrix(m._rows, m.Cols);
        for (var i = 0; i < m._rows; i++)
            for (var j = 0; j < m.Cols; j++)
                r[i, j] = m[i, j] * n;
        return r;
    }

    private static XLMatrix Add(XLMatrix m1, XLMatrix m2)
    {
        if (m1._rows != m2._rows || m1.Cols != m2.Cols)
            throw new ArgumentException("Matrices must have the same dimensions!");
        var r = new XLMatrix(m1._rows, m1.Cols);
        for (var i = 0; i < r._rows; i++)
            for (var j = 0; j < r.Cols; j++)
                r[i, j] = m1[i, j] + m2[i, j];
        return r;
    }

    public static string NormalizeMatrixString(string matStr) // From Andy - thank you! :)
    {
        // Remove any multiple spaces
        while (matStr.Contains("  "))
            matStr = matStr.Replace("  ", " ");

        // Remove any spaces before or after newlines
        matStr = matStr.Replace(" \r\n", "\r\n");
        matStr = matStr.Replace("\r\n ", "\r\n");

        // If the data ends in a newline, remove the trailing newline.
        // Make it easier by first replacing \r\n’s with |’s then
        // restore the |’s with \r\n’s
        matStr = matStr.Replace("\r\n", "|");
        while (matStr.LastIndexOf('|') == matStr.Length - 1)
            matStr = matStr[..^1];

        matStr = matStr.Replace("|", "\r\n");
        return matStr;
    }

    public static XLMatrix operator -(XLMatrix m)
    {
        return Multiply(-1, m);
    }

    public static XLMatrix operator +(XLMatrix m1, XLMatrix m2)
    {
        return Add(m1, m2);
    }

    public static XLMatrix operator -(XLMatrix m1, XLMatrix m2)
    {
        return Add(m1, -m2);
    }

    public static XLMatrix operator *(XLMatrix m1, XLMatrix m2)
    {
        return StrassenMultiply(m1, m2);
    }

    public static XLMatrix operator *(double n, XLMatrix m)
    {
        return Multiply(n, m);
    }
}
