import os
D = os.path.dirname(os.path.abspath(__file__))
p = os.path.join(D, "../../../cshell/collect.py")
s = open(p).read()
helpers = open(os.path.join(D, "_new_helpers.py")).read()
newc = open(os.path.join(D, "_new_collect.py")).read()
a = s.index("class CellData:")
b = s.index("def raw_responses(")
s = s[:a] + helpers + "\n\n" + s[b:]
a = s.index("def collect(")
b = s.index("def summary_text(")
s = s[:a] + newc.rstrip() + "\n\n\n" + s[b:]
open(p, "w").write(s)
print("spliced")
