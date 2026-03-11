using System;
using static ClosedXML.Excel.XLProtectionAlgorithm;

using System;
using static XLibur.Excel.XLProtectionAlgorithm;

namespace XLibur.Excel;

internal class XLSheetProtection : IXLSheetProtection
{
    public XLSheetProtection(Algorithm algorithm)
    {
        Algorithm = algorithm;
        AllowedElements = XLSheetProtectionElements.SelectEverything;
    }

    public Algorithm Algorithm { get; internal set; }
    public XLSheetProtectionElements AllowedElements { get; set; }

    public bool IsPasswordProtected => IsProtected && !string.IsNullOrEmpty(PasswordHash);
    public bool IsProtected { get; internal set; }

    internal string Base64EncodedSalt { get; set; } = string.Empty;
    internal string PasswordHash { get; set; } = string.Empty;
    internal uint SpinCount { get; set; } = 100000;

    public IXLSheetProtection AllowElement(XLSheetProtectionElements element, bool allowed = true)
    {
        if (!allowed)
            return DisallowElement(element);

        AllowedElements |= element;
        return this;
    }

    public IXLSheetProtection AllowEverything()
    {
        return AllowElement(XLSheetProtectionElements.Everything);
    }

    public IXLSheetProtection AllowNone()
    {
        AllowedElements = XLSheetProtectionElements.None;
        return this;
    }

    public object Clone()
    {
        return new XLSheetProtection(Algorithm)
        {
            IsProtected = IsProtected,
            PasswordHash = PasswordHash,
            SpinCount = SpinCount,
            Base64EncodedSalt = Base64EncodedSalt,
            AllowedElements = AllowedElements
        };
    }

    public XLSheetProtection CopyFrom(IXLElementProtection<XLSheetProtectionElements> sheetProtection)
    {
        if (sheetProtection is XLSheetProtection xlSheetProtection)
        {
            IsProtected = xlSheetProtection.IsProtected;
            Algorithm = xlSheetProtection.Algorithm;
            PasswordHash = xlSheetProtection.PasswordHash;
            SpinCount = xlSheetProtection.SpinCount;
            Base64EncodedSalt = xlSheetProtection.Base64EncodedSalt;
            AllowedElements = xlSheetProtection.AllowedElements;
        }
        return this;
    }

    public IXLSheetProtection DisallowElement(XLSheetProtectionElements element)
    {
        AllowedElements &= ~element;
        return this;
    }

    public IXLSheetProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        return Protect(string.Empty, algorithm);
    }

    public IXLSheetProtection Protect(XLSheetProtectionElements allowedElements)
        => Protect(string.Empty, DefaultProtectionAlgorithm, allowedElements);

    public IXLSheetProtection Protect(Algorithm algorithm, XLSheetProtectionElements allowedElements)
        => Protect(string.Empty, algorithm, allowedElements);

    public IXLSheetProtection Protect(string password, Algorithm algorithm = DefaultProtectionAlgorithm, XLSheetProtectionElements allowedElements = XLSheetProtectionElements.SelectEverything)
    {
        if (IsProtected)
        {
            throw new InvalidOperationException("The worksheet is already protected");
        }

        IsProtected = true;

        password ??= "";

        Algorithm = algorithm;
        Base64EncodedSalt = Utils.CryptographicAlgorithms.GenerateNewSalt(Algorithm);
        PasswordHash = Utils.CryptographicAlgorithms.GetPasswordHash(Algorithm, password, Base64EncodedSalt, SpinCount);

        AllowedElements = allowedElements;

        return this;
    }

    public IXLSheetProtection Unprotect()
    {
        return Unprotect(string.Empty);
    }

    public IXLSheetProtection Unprotect(string password)
    {
        if (IsProtected)
        {
            if (PasswordHash.Length > 0 && string.IsNullOrEmpty(password))
                throw new InvalidOperationException("The worksheet is password protected");

            var hash = Utils.CryptographicAlgorithms.GetPasswordHash(Algorithm, password, Base64EncodedSalt, SpinCount);
            if (hash != PasswordHash)
                throw new ArgumentException("Invalid password");
            IsProtected = false;
            PasswordHash = string.Empty;
            Base64EncodedSalt = string.Empty;
        }

        return this;
    }

    #region IXLProtectable interface

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowElement(XLSheetProtectionElements element, bool allowed) => AllowElement(element, allowed);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowEverything() => AllowEverything();

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.AllowNone() => AllowNone();

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.CopyFrom(IXLElementProtection<XLSheetProtectionElements> protectable) => CopyFrom(protectable);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.DisallowElement(XLSheetProtectionElements element) => DisallowElement(element);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Protect(Algorithm algorithm) => Protect(algorithm);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Protect(string password, Algorithm algorithm) => Protect(password, algorithm);

    IXLSheetProtection IXLSheetProtection.Protect(XLSheetProtectionElements allowedElements) => Protect(allowedElements);

    IXLSheetProtection IXLSheetProtection.Protect(Algorithm algorithm, XLSheetProtectionElements allowedElements) => Protect(algorithm, allowedElements);

    IXLSheetProtection IXLSheetProtection.Protect(string password, Algorithm algorithm, XLSheetProtectionElements allowedElements) => Protect(password, algorithm, allowedElements);

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Unprotect() => Unprotect();

    IXLElementProtection<XLSheetProtectionElements> IXLElementProtection<XLSheetProtectionElements>.Unprotect(string password) => Unprotect(password);

    #endregion IXLProtectable interface
}
