"""
Runs the generated Excel macro on a copy of the BLANK v1 template in headless
LibreOffice (VBA compatibility mode) and saves the result, so the macro route
can be compared with build.py without Excel.

Usage: python3 test_vba_libreoffice.py <macro.bas built with --test> <blank template.xlsx> <out.xlsx>
"""
import os
import re
import shutil
import subprocess
import sys
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def main(bas, template, out):
    work = os.path.dirname(os.path.abspath(out))
    src = os.path.join(work, "vba_input.xlsx")
    shutil.copy(template, src)
    pipe = f"vbatest{os.getpid()}"
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
                             f"--accept=pipe,name={pipe};urp;", f"-env:UserInstallation=file://{work}/lo_profile"])
    try:
        ctx = None
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(60):
            try:
                ctx = resolver.resolve(f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(uno.systemPathToFileUrl(src), "_blank", 0,
                                           (prop("Hidden", True), prop("MacroExecutionMode", 4)))
        time.sleep(8)                      # let the document finish loading before adding macros
        code = open(bas, encoding="ascii").read().replace("\r\n", "\n")
        lib = doc.BasicLibraries.getByName("Standard") if doc.BasicLibraries.hasByName("Standard") \
            else doc.BasicLibraries.createLibrary("Standard")
        # LibreOffice sometimes skips the first macro calls after a library is created: warm up first
        lib.insertByName("WarmUp", 'Option VBASupport 1\nPublic Function Ping() As String\n Ping = "ok"\nEnd Function\n')
        sp = doc.getScriptProvider()
        for _ in range(5):
            if sp.getScript("vnd.sun.star.script:Standard.WarmUp.Ping?language=Basic&location=document").invoke((), (), ())[0] == "ok":
                break
            time.sleep(1)
        # LibreOffice Basic modules are limited to about 64 KB of text (Excel has no such limit), so for this
        # test the data procedures go into extra modules; the helpers they call are made Public to allow that.
        split = code.index("' ====================================================== DATA")
        logic, data = code[:split], code[split:]
        for name in ("AddMD", "AddWT", "AddCC", "AddST"):
            logic = logic.replace(f"Private Sub {name}(", f"Public Sub {name}(")
        procs = re.split(r"\n(?=Private Sub Data)", data)
        mods, cur = [], ""
        for pr in procs:
            pr = pr.replace("Private Sub Data", "Public Sub Data", 1)
            if len(cur) + len(pr) > 55000:
                mods.append(cur)
                cur = ""
            cur += pr + "\n"
        mods.append(cur)
        lib.insertByName("Euro2028Build", "Option VBASupport 1\n" + logic)
        for i, m in enumerate(mods):
            lib.insertByName(f"Euro2028Data{i + 1}", "Option VBASupport 1\n" + m)
        script = sp.getScript(
            "vnd.sun.star.script:Standard.Euro2028Build.TestBuildNoPrompts?language=Basic&location=document")
        t0 = time.time()
        print("result:", script.invoke((), (), ())[0])
        print(f"macro ran in {time.time() - t0:.0f}s")
        doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(out)), (prop("FilterName", "Calc MS Excel 2007 XML"),))
        doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        proc.terminate()
    print(out)


if __name__ == "__main__":
    main(*sys.argv[1:4])
