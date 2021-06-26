using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelRead.Common
{
    public static class Extensions
    {
        public static bool ValidationTCNO(this string tcNo)
        {
            if (string.IsNullOrEmpty(tcNo))
                return false;

            Int64 tc = 0;

            bool isNumber = Int64.TryParse(tcNo, out tc);
            bool hasElevenChar = tcNo.Length == 11;
            bool isNotStartWithZero = tcNo.StartsWith("0") == false;


            // şartlar sağlanmıyorsa devam etme
            if (isNumber == false || hasElevenChar == false || isNotStartWithZero == false)
            {
                return false;
            }


            bool isValid = false;

            #region compute sums
            var first10NumberSum = 0;
            var evenNumbersSum = 0;
            var oddNumbersSum = 0;
            for (int i = 0; i < tcNo.Length; i++)
            {
                int n = (int)Char.GetNumericValue(tcNo[i]); // değerini al

                var isLastNo = i == 10;
                if (!isLastNo)
                {
                    first10NumberSum += n;
                }

                var isEven = i % 2 == 0;
                if (isEven && i <= 8) // i = 0,2,4,6,8 
                {
                    evenNumbersSum += n;
                }

                var isOdd = i % 2 != 0;
                if (isOdd && i < 8) // i = 1,3,5,7
                {
                    oddNumbersSum += n;
                }
            }
            #endregion

            var numberBeforeLastNumber = (int)Char.GetNumericValue(tcNo[9]);
            var lastNumber = (int)Char.GetNumericValue(tcNo[10]);

            // ilk 10 rakamın toplamının birler basamağı, son rakama eşittir
            bool isFirst10NumberSumEqualLastNumber = ((first10NumberSum % 10) == lastNumber);

            // ( çift haneler * 7 + tek haneler * 9 ) mod 10 = sondan bir önceki
            bool isRuleTrue = ((evenNumbersSum * 7 + oddNumbersSum * 9) % 10) == numberBeforeLastNumber;

            // çift haneler toplamının 8 katının birler basamağı(mod10) son rakama eşittir
            bool isEvenRuleTrue = (evenNumbersSum * 8) % 10 == lastNumber;

            // extra
            bool rule = (evenNumbersSum * 7 - oddNumbersSum) % 10 == numberBeforeLastNumber;
            bool isLastNumberEven = lastNumber % 2 == 0;

            if (isFirst10NumberSumEqualLastNumber && isEvenRuleTrue && isRuleTrue && rule && isLastNumberEven)
            {
                isValid = true;
            }

            return isValid;

        }
        public static string Tostring(this object value)
        {
            if (value == null)
                return string.Empty;

            if (value is Double)
                return ((Double)value).ToString(CultureInfo.InvariantCulture);

            return value.ToString();
        }
    }
}
