using Microsoft.Identity.Client;
using System;
using Gtk;

namespace MSTeamsHistory
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            Application.Init();
            MainWindow win = new MainWindow();
            win.Show();
            Application.Run();
        }

    }
}
