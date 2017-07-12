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
            Assert.True(HasChildWindow("explorer"));
        }

        [Fact]
        public void FindChildWindowFails() {
            Assert.False(HasChildWindow("xyzzy"));
        }

        private static IEnumerable<IntPtr> GetMainWindowHandles(string processName) {
            var processes = Process.GetProcessesByName(processName);
            return processes.Select(instance => instance.MainWindowHandle);
        }

        private bool HasChildWindow(string processName) {
            var mainWindows = GetMainWindowHandles(processName);
            return mainWindows.Any(mainWindow => finder.TryFindChildWindow(mainWindow, out _));
        }
    }
}
