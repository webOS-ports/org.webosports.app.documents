/**
	Fits an open ModalDialog to the window: its scroller as tall as what it
	holds, or as the window leaves with the rest of the dialog around it, and
	the dialog centred again. A dialog is centred on the size it has when it
	opens, so fill it in first, open it, then fit it; and fit it again when
	the window changes (a notification, the keyboard).

		dialog.validateComponents();   // a lazy dialog's controls, before filling them
		... fill it in ...
		scroller.applyStyle("height", null);
		dialog.openAtCenter();
		DialogFit.fit(dialog, scroller, content);
*/
var DialogFit = {
	//* Kept clear between the dialog and the window's edges
	MARGIN: 12,

	fit: function(inDialog, inScroller, inContent) {
		var d = inDialog.hasNode(), sc = inScroller.hasNode(), c = inContent.hasNode();
		if (!inDialog.isOpen || !d || !sc || !c) {
			return;
		}
		var chrome = d.offsetHeight - sc.offsetHeight;
		var room = Math.max(60, window.innerHeight - 2 * this.MARGIN - chrome);
		inScroller.applyStyle("height", Math.min(c.offsetHeight, room) + "px");
		inScroller.resized();
		inDialog.applyBoundsInfo();
	}
};
