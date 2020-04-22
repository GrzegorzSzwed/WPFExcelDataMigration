using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TextMatchCalculation
{
    public static class Levenstein
    {
        private static char[] _x;
        private static char[] _y;

        public static int Distance(string firstString, string SecondString)
        {
            _x = firstString.ToCharArray();
            _y = SecondString.ToCharArray();
            return CountDistance();
        }

        public static int NettoDistance(string firstString, string SecondString)
        {
            int minimum = System.Math.Min(firstString.Length, SecondString.Length);

            //count distance
            int distance = Distance(firstString, SecondString);

            return distance - minimum;
        }

        public static decimal Percent(string firstString, string SecondString)
        {
            //count requirements
            int maximum = System.Math.Max(firstString.Length, SecondString.Length);
            int minimum = System.Math.Min(firstString.Length, SecondString.Length);
            int gap = maximum - minimum;

            //count distance
            int distance = Distance(firstString, SecondString);

            //count percent
            if (distance >= gap && distance <= maximum)
            {
                decimal percent = 0;
                if (maximum != minimum)
                {
                    percent = 100 - System.Math.Round(
                        (decimal)(distance - gap) / (decimal)(maximum - gap),
                        4,
                        MidpointRounding.ToEven) * 100;
                }
                else
                {
                    percent = 100 - System.Math.Round(
                        (decimal)(distance / maximum),
                        4,
                        MidpointRounding.ToEven) * 100;
                }

                return percent;
            }
            else
            {
                throw new NotImplementedException();
            }
        }

        private static int CountDistance()
        {
            int[,] matrix = new int[_x.Length, _y.Length];

            //fill first row
            for (int i = 1; i < _x.Length; i++)
                matrix[i, 0] = i;

            //fill second row
            for (int j = 1; j < _y.Length; j++)
                matrix[0, j] = j;

            //fill matrix
            int substitutionCost = 0;

            for (int i = 1; i < _x.Length; i++)
            {
                for (int j = 1; j < _y.Length; j++)
                {
                    if (_y[j - 1] == _x[i - 1])
                        substitutionCost = 0;
                    else
                        substitutionCost = 1;

                    matrix[i, j] = Min3(matrix[i, j - 1] + 1, matrix[i - 1, j] + 1, matrix[i - 1, j - 1] + substitutionCost);
                }

            }

            //return last no
            return matrix[_x.Length - 1, _y.Length - 1];
        }

        private static int Min3(int v1, int v2, int v3)
        {
            return System.Math.Min(System.Math.Min(v1, v2), v3);
        }
    }
}
