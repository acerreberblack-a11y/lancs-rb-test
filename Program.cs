using System;

namespace LandocsRobot
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            var application = new RobotApplication();
            application.Run();
        }
    }
}
