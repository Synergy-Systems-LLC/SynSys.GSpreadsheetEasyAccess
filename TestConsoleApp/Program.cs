﻿using System;
using System.IO;
using GetGoogleSheetDataAPI;

namespace TestConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var connector = new Connector();
            bool isConnectionSuccessful = connector.TryToCreateConnect(CredentialStream());

            if (!isConnectionSuccessful)
            {
                if (connector.Status == ConnectStatus.NotConnected)
                {
                    Console.WriteLine("Подключение отсутствует");
                }
                else if (connector.Status == ConnectStatus.AuthorizationTimedOut)
                {
                    Console.WriteLine("Время на подключение закончилось");
                }
                
                return;
            }

            Console.WriteLine("Подключились к Cloud App");

            //string url = "https://docs.google.com/spreadsheets/d/1dxPz9MEeJxfZkbAvZLE33YLIN5GaNS0Bvqzlp6rAiNk/edit#gid=0";

            //if (HttpManager.CheckURL(url))
            //{
            //    var sheet = connector.TryToCreateSheet(url);

            //    //Console.WriteLine(data.Values.Count.ToString());
            //}
            //else
            //{
            //    Console.WriteLine(HttpManager.Status);
            //}

            Console.ReadLine();
        }

        private static Stream CredentialStream()
        {
            return new MemoryStream(Properties.Resources.credentials);
        }
    }
}
