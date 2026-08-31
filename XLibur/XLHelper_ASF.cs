/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at
       http://www.apache.org/licenses/LICENSE-2.0
   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

using System;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace XLibur.Excel;

public static partial class XLHelper
{
    private const int MaxWorksheetNameCharsCount = 31;
    private static readonly char[] IllegalWorksheetCharacters = "\0\u0003:\\/?*[]".ToCharArray();

    /// <summary>
    /// Creates a valid sheet name, which conforms to the following rules.
    /// A sheet name:
    /// is never null,
    /// has a minimum length of 1 and
    /// maximum length of 31,
    /// doesn't contain special chars: (: 0x0000 0x0003 / \ ? * ] [).
    /// Sheet names must not begin or end with ' (apostrophe)
    /// </summary>
    /// <param name="nameProposal">can be any string, will be truncated if necessary, allowed to be null</param>
    /// <param name="replaceChar">the char to replace invalid characters.</param>
    /// <returns>a valid string, "empty" if too short, "null" if null</returns>
    // This method was ported and adapted from the POI project at https://github.com/apache/poi/blob/trunk/src/java/org/apache/poi/ss/util/WorkbookUtil.java
    public static string CreateSafeSheetName(string nameProposal, char replaceChar = ' ')
    {
        if (IllegalWorksheetCharacters.Contains(replaceChar) || replaceChar == '\'')
            throw new ArgumentException("Invalid replacement character.", nameof(replaceChar));

        if (nameProposal == null)
        {
            return "null";
        }
        if (nameProposal.Length < 1)
        {
            return "empty";
        }
        var length = Math.Min(MaxWorksheetNameCharsCount, nameProposal.Length);
        var shortenedName = nameProposal[..length];
        var result = new StringBuilder(shortenedName);
        for (var i = 0; i < length; i++)
        {
            var ch = result[i];
            if (IllegalWorksheetCharacters.Contains(result[i]))
                result[i] = replaceChar;

            if (ch == '\'' && (i == 0 || i == length - 1))
                result[i] = replaceChar;
        }
        return result.ToString();
    }

    /// <summary>
    /// Validates the name of the sheet.
    /// The character count MUST be greater than or equal to 1 and less than or equal to 31.
    /// The string MUST NOT contain any of the following characters:
    /// - 0x0000
    /// - 0x0003
    /// - colon (:)
    /// - backslash(\)
    /// - asterisk(*)
    /// - question mark(?)
    /// - forward slash(/)
    /// - opening square bracket([)
    /// - closing square bracket(])
    /// The string MUST NOT begin or end with the single quote(') character.
    /// </summary>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <exception cref="ArgumentException">
    /// </exception>
    ///
    // This method was ported from the POI project at https://github.com/apache/poi/blob/trunk/src/java/org/apache/poi/ss/util/WorkbookUtil.java
    public static void ValidateSheetName(string sheetName)
    {
        if (!TryValidateSheetName(sheetName, out var reason))
            throw new ArgumentException(reason);
    }

    /// <summary>
    /// Test a sheet name against the same rules as <see cref="ValidateSheetName"/>, reporting the
    /// reason rather than throwing.
    /// </summary>
    /// <remarks>
    /// The rules are the same; who is being told off is not. <see cref="ValidateSheetName"/>
    /// answers a caller who passed the name in and can correct it, so
    /// <see cref="ArgumentException"/> is right there. A name read out of a workbook has no such
    /// caller — the person loading the file never chose it — and reporting a bad *argument* names
    /// a parameter they never supplied. The load path uses this overload and raises its own
    /// rejection instead (D33).
    /// </remarks>
    /// <param name="sheetName">Name of the sheet.</param>
    /// <param name="reason">Why the name is invalid, or <c>null</c> when it is valid.</param>
    /// <returns><c>true</c> when the name is a legal sheet name.</returns>
    internal static bool TryValidateSheetName(string sheetName, [NotNullWhen(false)] out string? reason)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            reason = "sheetName must not be null or whitespace";
            return false;
        }

        var len = sheetName.Length;
        if (len is < 1 or > MaxWorksheetNameCharsCount)
        {
            reason = $"sheetName '{sheetName}' is invalid - character count MUST be greater than or equal to 1 and less than or equal to {MaxWorksheetNameCharsCount}";
            return false;
        }

        for (var i = 0; i < len; i++)
        {
            if (IllegalWorksheetCharacters.Contains(sheetName[i]))
            {
                reason = $"Invalid char ({sheetName[i]}) found at index ({i}) in sheet name '{sheetName}'";
                return false;
            }
        }

        if (sheetName[0] == '\'' || sheetName[len - 1] == '\'')
        {
            reason = $"Invalid sheet name '{sheetName}'. Sheet names must not begin or end with (').";
            return false;
        }

        reason = null;
        return true;
    }
}
