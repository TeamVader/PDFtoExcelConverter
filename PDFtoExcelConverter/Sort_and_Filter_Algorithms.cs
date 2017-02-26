using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PDFtoExcelConverter
{
    class Sort_and_Filter_Algorithms
    {
        public static bool findint(int search, int[] array)
        {
            for (int i = 0; i < array.Length; i++)
            {
                if (search == array[i])
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Lookup for String in array
        /// </summary>
        /// <param name="search"></param>
        /// <param name="array"></param>
        /// <returns></returns>
        private static bool findstring(string search, string[] array)
        {
            for (int i = 0; i < array.Length; i++)
            {
                if (search == array[i])
                {
                    return true;
                }
            }
            return false;
        }

        #region Clamps
        public static void Filter_Clamp_BMK(string[] bmk_names, string[] pdf_result)
        {
            int[] temp = new int[MainForm.size];
            int[] bmk_numbers = new int[MainForm.size];
            int dummy;
            int klemmearraypointer = 0;
            int finalarraypointer = 0;
            Regex klemme_number_regex = new Regex(@"[-+]?([0-9]{4,5})");
            Regex bmk_regex = new Regex(@"[-+]?([0-9]{1,4})[X][-+]?([0-9]*\.[0-9]{1,3}|[0-9]{1,3})");
            // find all numbers in string
            for (int j = 0; j < pdf_result.Length; j++)
            {
                foreach (Match match in klemme_number_regex.Matches(pdf_result[j]))
                {
                    if (Int32.TryParse(match.Value, out dummy))
                    {
                        if (dummy >= 999 && dummy <= 99999)
                        {
                            if (findint(dummy, temp) == false)
                            {
                                temp[klemmearraypointer] = dummy;

                                klemmearraypointer++;

                            }
                        }
                    }
                }

                foreach (Match match in bmk_regex.Matches(pdf_result[j]))
                {



                    if (findstring(match.Value.Replace("-", ""), bmk_names) == false)
                    {

                        bmk_names[finalarraypointer] = match.Value.Replace("-", "");
                        finalarraypointer++;

                    }
                }


            }

            klemmearraypointer = 0;
            //Sort array
            Array.Sort(temp);
            //find all valid clamp numbers 
            for (int i = 1; i < 50; i++)
            {
                for (int j = 100; j < 100000; j = j + 100)
                {
                    for (int k = 0; k < temp.Length; k++)
                    {
                        if (i == 1)
                        {
                            if ((j + i) == temp[k])
                            {
                                bmk_numbers[klemmearraypointer] = temp[k];
                                klemmearraypointer++;
                            }
                        }
                        else
                        {
                            if ((j + i) == temp[k])
                            {
                                if (findint(temp[k] - 1, bmk_numbers) == true)
                                {
                                    bmk_numbers[klemmearraypointer] = temp[k];
                                    klemmearraypointer++;
                                }
                            }
                        }
                    }


                }
            }
            Array.Sort(bmk_numbers);
            //add clamp numbers 
            for (int i = 0; i < bmk_numbers.Length; i++)
            {
                if (bmk_numbers[i] != 0)
                {
                    bmk_names[finalarraypointer] = bmk_numbers[i].ToString();
                    finalarraypointer++;
                }
            }


        }
        #endregion

        #region Cable
        public static void Filter_Cable_BMK(string[] bmk_names, string[] pdf_result)
        {
            string[] temp = new string[MainForm.size];

            Regex no_cable_letter = new Regex(@"[AaFfNnKkQqTtXx]");
            Regex number_regex = new Regex(@"[-+]?([0-9]{1,5})");
            Regex bmk_regex = new Regex(@"[-+]?([0-9]{1,4})[A-Z][-+]?([0-9]*\.[0-9]{1,3}|[0-9]{1,3})");
            List<string> templist = new List<string>();
            int arraypointer = 0;

            for (int j = 0; j < pdf_result.Length; j++)
            {

                //
                foreach (Match match in bmk_regex.Matches(pdf_result[j]))
                {



                    if (findstring(match.Value.Replace("-", ""), temp) == false)
                    {

                        temp[arraypointer] = match.Value.Replace("-", "");
                        arraypointer++;

                    }
                }

            }

            //Filter No Cable lettres
            for (int i = 0; i < temp.Length; i++)
            {
                if (!string.IsNullOrEmpty(temp[i]))
                {
                    if (no_cable_letter.IsMatch(temp[i]))
                    {
                        temp[i] = null;
                    }
                    else
                    {
                        templist.Add(temp[i]);
                    }
                }
            }

            arraypointer = 0;
            //sort cable letters
            #region Bubble sorting


            int number = 0;
            int next_number = 0;

            for (int write = 0; write < templist.Count; write++)
            {
                for (int sort = 0; sort < templist.Count - 1; sort++)
                {

                    Match match = number_regex.Match(templist[sort]);
                    if (match.Success)
                    {
                        if (Int32.TryParse(match.Value, out number))
                        {
                            match = number_regex.Match(templist[sort + 1]);
                            if (match.Success)
                            {
                                if (Int32.TryParse(match.Value, out next_number))
                                {
                                    if (number > next_number)
                                    {



                                        string stringtemp = templist[sort + 1];
                                        templist[sort + 1] = templist[sort];
                                        templist[sort] = stringtemp;

                                    }
                                }

                            }

                        }


                    }
                }
            }




            #endregion

            //Override array
            for (int i = 0; i < templist.Count; i++)
            {
                bmk_names[i] = templist[i];
            }
        }
        #endregion

        #region Bmk_Excel
        public static void Filter_Excel_BMK(string[] bmk_names, string[] pdf_result)
        {
            string[] temp = new string[MainForm.size];

            Regex number_regex = new Regex(@"[-+]?([0-9]{1,5})");
            Regex bmk_regex = new Regex(@"[-+]?([0-9]{1,4})[A-Z][-+]?([0-9]*\.[0-9]{1,3}|[0-9]{1,3})");
            List<string> templist = new List<string>();
            int arraypointer = 0;

            for (int j = 0; j < pdf_result.Length; j++)
            {

                //
                foreach (Match match in bmk_regex.Matches(pdf_result[j]))
                {



                    if (findstring(match.Value.Replace("-", ""), temp) == false)
                    {

                        temp[arraypointer] = match.Value.Replace("-", "");
                        arraypointer++;

                    }
                }

            }

            //write in list -> ev filtering for 
            for (int i = 0; i < temp.Length; i++)
            {
                if (!string.IsNullOrEmpty(temp[i]))
                {
                   
                        templist.Add(temp[i]);
                   
                }
            }

            //sort cable letters
            #region Bubble sorting


            int number = 0;
            int next_number = 0;

            for (int write = 0; write < templist.Count; write++)
            {
                for (int sort = 0; sort < templist.Count - 1; sort++)
                {

                    Match match = number_regex.Match(templist[sort]);
                    if (match.Success)
                    {
                        if (Int32.TryParse(match.Value, out number))
                        {
                            match = number_regex.Match(templist[sort + 1]);
                            if (match.Success)
                            {
                                if (Int32.TryParse(match.Value, out next_number))
                                {
                                    if (number > next_number)
                                    {



                                        string stringtemp = templist[sort + 1];
                                        templist[sort + 1] = templist[sort];
                                        templist[sort] = stringtemp;

                                    }
                                }

                            }

                        }


                    }
                }
            }




            #endregion

            //Override array
            for (int i = 0; i < templist.Count; i++)
            {
                bmk_names[i] = templist[i];
            }
        }
        #endregion



    }
}
