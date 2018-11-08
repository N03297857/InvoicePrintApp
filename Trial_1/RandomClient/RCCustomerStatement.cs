using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Trial_1.RandomClient
{
    class RCCustomerStatement
    {
        public string AccountNumber { get; set; }
        public string Description { get; set; }
        public int Price { get; set; }
        public int Tax { get; set; }

        public int DescriptionLine { get; private set; }

        private int GetLine(string aString, Font aFont, float aSize, float aWidth)
        {
            if (aFont == null || aSize <= 0) throw new ArgumentNullException("Font and size cannot be null or less/equal then 0.");

            if (String.IsNullOrEmpty(aString))
            {
                return 0;
            }

            BaseFont bf = aFont.GetCalculatedBaseFont(true);
            aString = aString.TrimStart(' ');
            aString = aString.TrimEnd(' ');
            string[] input = aString.Split(' ');
            int line = 1;

            string temp = "";
            for (int i = 0; i < input.Length; i++)
            {
                string currentWord = input[i];
                float currentWordSize = bf.GetWidthPoint(currentWord, 9);

                if (currentWordSize > aSize)
                {
                    for (int j = 0; j < currentWord.Length; j++)
                    {
                        temp += currentWord[j];
                        if (bf.GetWidthPoint(temp, 9) <= aWidth)
                        {
                            if (j == currentWord.Length - 1)
                            {
                                temp += " ";
                            }
                            continue;
                        }
                        else
                        {
                            line++;
                            if (j == currentWord.Length - 1)
                            {
                                temp = currentWord[j].ToString() + " ";
                            }
                            else
                            {
                                temp = currentWord[j].ToString();
                            }
                        }
                    }
                }
                else
                {
                    temp += input[i];
                    if (bf.GetWidthPoint(temp, 9) <= aWidth)
                    {
                        if (i == input.Length - 1)
                        {
                            temp += " ";
                        }
                        continue;
                    }
                    else
                    {
                        line++;
                        if (i == input.Length - 1)
                        {
                            temp = input[i];
                        }
                        else
                        {
                            temp = input[i] + " ";
                        }
                    }
                }
            }

            return line;
        }
        public void setDescription(string aDescription)
        {
            DescriptionLine = GetLine(aDescription, FontFactory.GetFont("Arial", 9), 9f, 255.555542f);
            Description = aDescription;
        }

        public override string ToString()
        {
            string str = "";

            str += "AccountNo-" + AccountNumber + ",";
            str += "Description-" + Description + ",";
            str += "Price-" + Price + ",";
            str += "Tax-" + Tax + ",";

            return str;
        }
    }
}
