@{
    # N name, F family, A argument kinds, T typical argument text per argument,
    # Min fewest arguments Excel accepts (default: all of A).
    # Kinds: num, int, text, bool, any, ref. Variadic functions list two arguments.
    # Left out on purpose: volatile or environment functions (TODAY, NOW, RAND, INFO, CELL),
    # functions that return arrays, and functions that need an _xlfn. prefix.
    Functions = @(
        # math
        @{ N = 'SIGN'; F = 'math'; A = @('num'); T = @('-5') }
        @{ N = 'SQRT'; F = 'math'; A = @('num'); T = @('16') }
        @{ N = 'EXP'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'LN'; F = 'math'; A = @('num'); T = @('10') }
        @{ N = 'LOG'; F = 'math'; A = @('num', 'num'); T = @('100', '10'); Min = 1 }
        @{ N = 'LOG10'; F = 'math'; A = @('num'); T = @('1000') }
        @{ N = 'CEILING'; F = 'math'; A = @('num', 'num'); T = @('2.5', '1') }
        @{ N = 'FLOOR'; F = 'math'; A = @('num', 'num'); T = @('2.5', '1') }
        @{ N = 'ROUNDUP'; F = 'math'; A = @('num', 'int'); T = @('2.345', '2') }
        @{ N = 'ROUNDDOWN'; F = 'math'; A = @('num', 'int'); T = @('2.345', '2') }
        @{ N = 'TRUNC'; F = 'math'; A = @('num', 'int'); T = @('8.9', '0'); Min = 1 }
        @{ N = 'FACT'; F = 'math'; A = @('int'); T = @('5') }
        @{ N = 'FACTDOUBLE'; F = 'math'; A = @('int'); T = @('7') }
        @{ N = 'COMBIN'; F = 'math'; A = @('int', 'int'); T = @('5', '2') }
        @{ N = 'GCD'; F = 'math'; A = @('int', 'int'); T = @('12', '18'); Min = 1 }
        @{ N = 'LCM'; F = 'math'; A = @('int', 'int'); T = @('4', '6'); Min = 1 }
        @{ N = 'QUOTIENT'; F = 'math'; A = @('num', 'num'); T = @('7', '2') }
        @{ N = 'EVEN'; F = 'math'; A = @('num'); T = @('3') }
        @{ N = 'ODD'; F = 'math'; A = @('num'); T = @('2') }
        @{ N = 'MROUND'; F = 'math'; A = @('num', 'num'); T = @('10', '3') }
        @{ N = 'PI'; F = 'math'; A = @(); T = @() }
        @{ N = 'SIN'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'COS'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'TAN'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'ASIN'; F = 'math'; A = @('num'); T = @('0.5') }
        @{ N = 'ACOS'; F = 'math'; A = @('num'); T = @('0.5') }
        @{ N = 'ATAN'; F = 'math'; A = @('num'); T = @('1') }
        @{ N = 'ATAN2'; F = 'math'; A = @('num', 'num'); T = @('1', '1') }
        @{ N = 'DEGREES'; F = 'math'; A = @('num'); T = @('3.14159265358979') }
        @{ N = 'RADIANS'; F = 'math'; A = @('num'); T = @('180') }
        @{ N = 'ROMAN'; F = 'math'; A = @('int', 'any'); T = @('1999', '0'); Min = 1 }
        @{ N = 'MAX'; F = 'math'; A = @('any', 'any'); T = @('3', '7'); Min = 1 }
        @{ N = 'MIN'; F = 'math'; A = @('any', 'any'); T = @('3', '7'); Min = 1 }
        @{ N = 'SUMSQ'; F = 'math'; A = @('any', 'any'); T = @('3', '4'); Min = 1 }
        @{ N = 'SUMPRODUCT'; F = 'math'; A = @('ref', 'ref'); T = @('A1:A3', 'B1:B3'); Min = 1 }

        # text
        @{ N = 'CONCATENATE'; F = 'text'; A = @('text', 'text'); T = @('"a"', '"b"'); Min = 1 }
        @{ N = 'EXACT'; F = 'text'; A = @('text', 'text'); T = @('"a"', '"A"') }
        @{ N = 'LOWER'; F = 'text'; A = @('text'); T = @('"AbC"') }
        @{ N = 'UPPER'; F = 'text'; A = @('text'); T = @('"AbC"') }
        @{ N = 'PROPER'; F = 'text'; A = @('text'); T = @('"hello wORLD"') }
        @{ N = 'RIGHT'; F = 'text'; A = @('text', 'int'); T = @('"hello"', '2'); Min = 1 }
        @{ N = 'REPT'; F = 'text'; A = @('text', 'int'); T = @('"ab"', '3') }
        @{ N = 'REPLACE'; F = 'text'; A = @('text', 'int', 'int', 'text'); T = @('"abcdef"', '2', '3', '"X"') }
        @{ N = 'SEARCH'; F = 'text'; A = @('text', 'text', 'int'); T = @('"l"', '"hello"', '1'); Min = 2 }
        @{ N = 'CHAR'; F = 'text'; A = @('int'); T = @('65') }
        @{ N = 'CODE'; F = 'text'; A = @('text'); T = @('"A"') }
        @{ N = 'CLEAN'; F = 'text'; A = @('text'); T = @('"abc"') }
        @{ N = 'T'; F = 'text'; A = @('any'); T = @('"abc"') }

        # logical
        @{ N = 'NOT'; F = 'logical'; A = @('any'); T = @('TRUE') }
        @{ N = 'OR'; F = 'logical'; A = @('any', 'any'); T = @('FALSE', 'TRUE'); Min = 1 }
        @{ N = 'TRUE'; F = 'logical'; A = @(); T = @() }
        @{ N = 'FALSE'; F = 'logical'; A = @(); T = @() }

        # information
        @{ N = 'ISBLANK'; F = 'information'; A = @('any'); T = @('A1') }
        @{ N = 'ISERR'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ISERROR'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ISEVEN'; F = 'information'; A = @('num'); T = @('4') }
        @{ N = 'ISODD'; F = 'information'; A = @('num'); T = @('3') }
        @{ N = 'ISLOGICAL'; F = 'information'; A = @('any'); T = @('TRUE') }
        @{ N = 'ISNA'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ISNONTEXT'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ISTEXT'; F = 'information'; A = @('any'); T = @('"a"') }
        @{ N = 'ISREF'; F = 'information'; A = @('any'); T = @('A1') }
        @{ N = 'NA'; F = 'information'; A = @(); T = @() }
        @{ N = 'TYPE'; F = 'information'; A = @('any'); T = @('1') }
        @{ N = 'ERROR.TYPE'; F = 'information'; A = @('any'); T = @('#N/A') }
    )
}
