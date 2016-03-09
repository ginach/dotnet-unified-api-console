﻿#region

using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
#endregion

namespace MicrosoftGraphSampleConsole
{
    internal class Helper
    {
        /// <summary>
        ///     Returns a random string of upto 32 characters.
        /// </summary>
        /// <returns>String of upto 32 characters.</returns>
        public static string GetRandomString(int length = 32)
        {
            //because GUID can't be longer than 32
            return Guid.NewGuid().ToString("N").Substring(0, length > 32 ? 32 : length);
        }
    }
}