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
		AppKit.NSButton conservative_rounding { get; set; }

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

		[Outlet]
		AppKit.NSButton use_analytical_predictions_checkbox { get; set; }

		[Outlet]
		AppKit.NSTextField user_discount_rate { get; set; }

		[Outlet]
		AppKit.NSImageView user_growth_rate_icon { get; set; }

		[Outlet]
		AppKit.NSTextField user_growth_rate_textbox { get; set; }

		[Action ("btn_calculate:")]
		partial void btn_calculate (Foundation.NSObject sender);

		[Action ("conservative_rounds_checkbox:")]
		partial void conservative_rounds_checkbox (Foundation.NSObject sender);

		[Action ("perpertual_growth_rate_textbox:")]
		partial void perpertual_growth_rate_textbox (Foundation.NSObject sender);

		[Action ("stock_textbox:")]
		partial void stock_textbox (Foundation.NSObject sender);

		[Action ("user_discount_rate_textbox:")]
		partial void user_discount_rate_textbox (Foundation.NSObject sender);

		[Action ("user_growth_rate_textbox_action:")]
		partial void user_growth_rate_textbox_action (Foundation.NSObject sender);
		
		void ReleaseDesignerOutlets ()
		{
			if (btn_generate_click != null) {
				btn_generate_click.Dispose ();
				btn_generate_click = null;
			}

			if (conservative_rounding != null) {
				conservative_rounding.Dispose ();
				conservative_rounding = null;
			}

			if (perpertual_growth_rate != null) {
				perpertual_growth_rate.Dispose ();
				perpertual_growth_rate = null;
			}

			if (perpertual_growth_rate_icon != null) {
				perpertual_growth_rate_icon.Dispose ();
				perpertual_growth_rate_icon = null;
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

			if (use_analytical_predictions_checkbox != null) {
				use_analytical_predictions_checkbox.Dispose ();
				use_analytical_predictions_checkbox = null;
			}

			if (user_discount_rate != null) {
				user_discount_rate.Dispose ();
				user_discount_rate = null;
			}

			if (user_growth_rate_textbox != null) {
				user_growth_rate_textbox.Dispose ();
				user_growth_rate_textbox = null;
			}

			if (user_growth_rate_icon != null) {
				user_growth_rate_icon.Dispose ();
				user_growth_rate_icon = null;
			}
		}
	}
}
