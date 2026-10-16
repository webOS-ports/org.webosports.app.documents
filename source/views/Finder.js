/**
	Find in a rendered document: marks every match in the text under a
	node, ignoring case and accents ("cafe" finds "Café").

		var f = new DocFinder(node);
		var count = f.search("cafe");   // marks them: <mark class="find-hit">
		f.select(2);                    // the third is the current one (find-current); returns its node
		f.clear();                      // takes the marks out again

	A match must lie inside one run of text: a word split across two
	differently styled spans is not found.
*/
function DocFinder(inRoot) {
	this.root = inRoot;
	this.marks = [];
	this.current = -1;
}

//* Lower case, accents dropped; one output string per input character (so offsets map back).
DocFinder.fold = function(inChar) {
	var c = inChar;
	if (c.normalize) {
		c = c.normalize("NFD").replace(/[̀-ͯ]/g, "");
	}
	return c.toLowerCase();
};

DocFinder.foldString = function(inText) {
	var s = "";
	for (var i = 0; i < inText.length; i++) {
		s += DocFinder.fold(inText.charAt(i));
	}
	return s;
};

DocFinder.prototype.clear = function() {
	for (var i = 0; i < this.marks.length; i++) {
		var m = this.marks[i];
		var parent = m.parentNode;
		if (!parent) {
			continue;
		}
		while (m.firstChild) {
			parent.insertBefore(m.firstChild, m);
		}
		parent.removeChild(m);
		parent.normalize();
	}
	this.marks = [];
	this.current = -1;
};

DocFinder.prototype.textNodes = function() {
	var out = [];
	var walker = document.createTreeWalker(this.root, NodeFilter.SHOW_TEXT, {
		acceptNode: function(inNode) {
			var p = inNode.parentNode;
			if (!p || /^(SCRIPT|STYLE|TEXTAREA)$/.test(p.nodeName)) {
				return NodeFilter.FILTER_REJECT;
			}
			return inNode.nodeValue ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
		}
	}, false);
	var n;
	while ((n = walker.nextNode())) {
		out.push(n);
	}
	return out;
};

//* Marks every match; returns how many (at most inMax, 5000 by default).
DocFinder.prototype.search = function(inQuery, inMax) {
	this.clear();
	var q = DocFinder.foldString(String(inQuery || "").replace(/^\s+|\s+$/g, ""));
	if (!q) {
		return 0;
	}
	var max = inMax || 5000;
	var nodes = this.textNodes();
	for (var i = 0; i < nodes.length && this.marks.length < max; i++) {
		var node = nodes[i];
		var text = node.nodeValue;
		// folded text and, per folded character, the original index
		var folded = "", map = [];
		for (var k = 0; k < text.length; k++) {
			var f = DocFinder.fold(text.charAt(k));
			for (var j = 0; j < f.length; j++) {
				map.push(k);
			}
			folded += f;
		}
		map.push(text.length);
		var hits = [];
		var from = 0, at;
		while ((at = folded.indexOf(q, from)) >= 0 && this.marks.length + hits.length < max) {
			hits.push([map[at], map[at + q.length - 1] + 1]);
			from = at + q.length;
		}
		// Wrap from the end, so earlier offsets stay valid.
		var wrapped = [];
		for (var h = hits.length - 1; h >= 0; h--) {
			var start = hits[h][0], end = hits[h][1];
			var target = node;
			if (end < target.nodeValue.length) {
				target.splitText(end);
			}
			var matchNode = start > 0 ? target.splitText(start) : target;
			var mark = document.createElement("mark");
			mark.className = "find-hit";
			matchNode.parentNode.insertBefore(mark, matchNode);
			mark.appendChild(matchNode);
			wrapped.unshift(mark);
		}
		for (var w = 0; w < wrapped.length; w++) {
			this.marks.push(wrapped[w]);
		}
	}
	return this.marks.length;
};

//* Makes match inIndex the current one; returns its mark element.
DocFinder.prototype.select = function(inIndex) {
	if (!this.marks.length) {
		return null;
	}
	if (this.current >= 0 && this.marks[this.current]) {
		this.marks[this.current].className = "find-hit";
	}
	var n = this.marks.length;
	this.current = ((inIndex % n) + n) % n;
	var m = this.marks[this.current];
	m.className = "find-hit find-current";
	return m;
};
