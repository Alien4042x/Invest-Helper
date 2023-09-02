// WARNING
//
// This file has been generated automatically by Visual Studio to store outlets and
// actions made in the UI designer. If it is removed, they will be lost.
// Manual changes to this file may not be handled correctly.
//
using Foundation;
using System.CodeDom.Compiler;

namespace InvestHelper
{
	[Register ("ViewController")]
	partial class ViewController
	{
		[Outlet]
		AppKit.NSButton btn_generate_click { get; set; }

		[Outlet]
		AppKit.NSTextField discount_rate { get; set; }

		[Outlet]
		AppKit.NSImageView discount_rate_icon { get; set; }

		[Outlet]
		AppKit.NSTextField growth_rate { get; set; }

		[Outlet]
		AppKit.NSImageView growth_rate_icon { get; set; }

		[Outlet]
		AppKit.NSTextField perpertual_growth_rate { get; set; }

		[Outlet]
		AppKit.NSImageView perpertual_growth_rate_icon { get; set; }

		[Outlet]
		AppKit.NSProgressIndicator progress_indicator { get; set; }

		[Outlet]
		AppKit.NSTextField stock { get; set; }

		[Outlet]
		AppKit.NSImageView stock_icon { get; set; }

		[Action ("btn_calculate:")]
		partial void btn_calculate (Foundation.NSObject sender);

		[Action ("discount_rate_textbox:")]
		partial void discount_rate_textbox (Foundation.NSObject sender);

		[Action ("growth_rate_textbox:")]
		partial void growth_rate_textbox (Foundation.NSObject sender);

		[Action ("perpertual_growth_rate_textbox:")]
		partial void perpertual_growth_rate_textbox (Foundation.NSObject sender);

		[Action ("stock_textbox:")]
		partial void stock_textbox (Foundation.NSObject sender);
		
		void ReleaseDesignerOutlets ()
		{
			if (btn_generate_click != null) {
				btn_generate_click.Dispose ();
				btn_generate_click = null;
			}

			if (discount_rate != null) {
				discount_rate.Dispose ();
				discount_rate = null;
			}

			if (growth_rate != null) {
				growth_rate.Dispose ();
				growth_rate = null;
			}

			if (perpertual_growth_rate != null) {
				perpertual_growth_rate.Dispose ();
				perpertual_growth_rate = null;
			}

			if (progress_indicator != null) {
				progress_indicator.Dispose ();
				progress_indicator = null;
			}

			if (stock != null) {
				stock.Dispose ();
				stock = null;
			}

			if (stock_icon != null) {
				stock_icon.Dispose ();
				stock_icon = null;
			}

			if (growth_rate_icon != null) {
				growth_rate_icon.Dispose ();
				growth_rate_icon = null;
			}

			if (perpertual_growth_rate_icon != null) {
				perpertual_growth_rate_icon.Dispose ();
				perpertual_growth_rate_icon = null;
			}

			if (discount_rate_icon != null) {
				discount_rate_icon.Dispose ();
				discount_rate_icon = null;
			}
		}
	}
}
