using AppKit;
using Foundation;

namespace InvestHelper
{
	[Register ("AppDelegate")]
	public class AppDelegate : NSApplicationDelegate
	{
		public AppDelegate ()
		{
		}

		public override void DidFinishLaunching (NSNotification notification)
		{
			// Insert code here to initialize your application
		}

		public override void WillTerminate (NSNotification notification)
		{
			// Insert code here to tear down your application
		}

        [Export("applicationSupportsSecureRestorableState:")]
        public bool ApplicationSupportsSecureRestorableState(NSApplication app)
        {
            return true;
        }

        public override bool ApplicationShouldTerminateAfterLastWindowClosed(NSApplication sender)
        {
            return true;
        }
    }
}

