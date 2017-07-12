using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelSerialPortListener;
using Xunit;

namespace ExcelSerialPortListener.Tests
{
    public sealed class ChildWindowFinderTests {
        private readonly ChildWindowFinder finder;

        public ChildWindowFinderTests() {
            const string windowClass = "TrayNotifyWnd";
            finder = ChildWindowFinder.FindWindowClass(windowClass);
        }

        [Fact]
        public void FindChildWindow() {
            var explorers = Process.GetProcessesByName("explorer");
            var mainWindows = explorers.Select(instance => instance.MainWindowHandle);
            bool succeeded = false;
            foreach (var mainWindow in mainWindows) {
                if (finder.TryFindChildWindow(mainWindow, out _)) {
                    succeeded = true;
                    break;
                }
            }
            Assert.True(succeeded);
        }

    }
}
