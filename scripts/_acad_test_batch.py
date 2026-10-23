import sys
import pythoncom
from win32com.client import GetObject

sys.path.insert(0, r"c:\repos\microstation-vba-project\scripts")
from _acad_batch_sample_fl import sample_many, OUT_LISP

ALL = [
    "13D80A", "13D2EA", "13D425", "13D42D", "13D2EE", "13D7BC", "13D7C4", "13D7CC",
    "13D7BF", "13CEC7", "13CED4", "13CED2", "13D25C", "13D263", "13D265", "13D25E",
    "13D282", "13D7DC", "13D7DB", "13D7E0", "13D7E8", "13D822", "13D824",
]

def lisp_len(handles):
    step = 0.4
    parts = [
        "(progn (vl-load-com)",
        f'(setq f (open "{OUT_LISP}" "w"))',
        f"(setq step {step:.4f})",
    ]
    for h in handles:
        parts.append(
            f'(progn (setq e (handent "{h}")) '
            f"(if e (progn "
            f"(setq len (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e))) "
            f"(setq d 0.0) "
            f"(while (<= d (+ len 1e-6)) "
            f"(setq p (vlax-curve-getPointAtDist e (min d len))) "
            f'(if p (write-line (strcat "{h}," (rtos (car p) 2 6) "," '
            f'(rtos (cadr p) 2 6)) f)) '
            f"(setq d (+ d step))) "
            f"))) "
        )
    parts.append('(write-line "DONE" f) (close f) (princ))')
    return len(" ".join(parts) + " ")


def main():
    pythoncom.CoInitialize()
    doc = GetObject(None, "AutoCAD.Application").ActiveDocument
    print("lisp chars all", lisp_len(ALL))
    d = sample_many(doc, ["13D80A", "13D2EA"], step=0.5, timeout=45)
    print("2-handle", {k: len(v) for k, v in d.items()})


if __name__ == "__main__":
    main()
