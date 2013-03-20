using System;

namespace CADProgrammingUtilities
{
    public class Incrementer
    {
        public static string ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        public static string alpha = "abcdefghijklmnopqrstuvwxyz";

        public static string IntToALPHA(int inputInt)
        {
            // Accepts an integer
            // Returns a string of upper alpha characters
            // A, B, C....Z, AA, AB, AC...AZ, BA, BB, BC .....ZZ, AAA, AAB, AAC, AAD
            string res = string.Empty;
            int iVal = inputInt;
            while (iVal >= 0)
            {
                int indx = (int)(iVal % 26);
                res = ALPHA.Substring(indx, 1) + res;
                iVal = iVal - indx;
                if (iVal == 0) iVal = -1;
                else iVal = (iVal / 26) - 1;
            }
            return res;
        }

        public static int ALPHAToInt(string inputString)
        {
            int res = 0;
            inputString = inputString.Trim();
            int len = inputString.Length;

            // start from the left
            // but opposite
            int ctr = 0;
            for (int i = len - 1; i > -1; i--)
            {
                int mxp = (int)Math.Pow(26, i);
                string ltr = inputString.Substring(ctr, 1);
                int nxov = ALPHA.IndexOf(ltr);
                res += mxp * (nxov + 1);
                ctr++;
            }
            return res;
        }

        public static string TickALPHA(string inputString)
        {
            return TickALPHA(inputString, true);
        }

        public static string TickALPHA(string inputString, bool Up)
        {
            string res = string.Empty;
            int curr = ALPHAToInt(inputString);
            if (!Up)
            {
                curr--;
                curr--;
            }
            res = IntToALPHA(curr);
            return res;
        }
    }
}