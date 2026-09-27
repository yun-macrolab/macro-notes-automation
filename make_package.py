"""Package only an explicit allowlist; exclude runtime, credentials and raw logs."""
from pathlib import Path
import json
import subprocess
import sys
import tempfile
import zipfile

BASE=Path(__file__).resolve().parent

if __name__=="__main__":
    files=[]
    for name in ("app.py","core.py","verify.py","build_plan.py","README.md","requirements.txt",
                 "Start-Demo.ps1","Stop-Demo.ps1","Start-Demo.bat","Stop-Demo.bat"):
        files.append(BASE/name)
    for folder,extensions in {"engine":{'.py'},"data":{'.json'},"static":{'.html','.css','.js'},
                              "docs":{'.md'},"exercises":{'.py'},"evidence":{'.json','.md','.png'}}.items():
        files.extend(p for p in (BASE/folder).iterdir() if p.is_file() and p.suffix in extensions)
    target=BASE/"portfolio-demo.zip"
    with zipfile.ZipFile(target,"w",zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(files):
            archive.write(path,"career-lab/"+path.relative_to(BASE).as_posix())
    with zipfile.ZipFile(target) as archive:
        assert archive.testzip() is None
        assert all("runtime" not in name and ".env" not in name for name in archive.namelist())
        with tempfile.TemporaryDirectory(prefix="macro_portable_") as tmp:
            root=Path(tmp).resolve()
            for name in archive.namelist():
                assert (root/name).resolve().is_relative_to(root)
            archive.extractall(root)
            result=subprocess.run([sys.executable,"-X","utf8","-c",
                "from pathlib import Path; from core import DemoStore; s=DemoStore(Path('runtime')); r=s.run(); assert r['news_count']==8; print('PORTABLE_OK')"],
                cwd=root/"career-lab",capture_output=True,text=True,encoding="utf-8",timeout=60)
            if result.returncode:
                raise RuntimeError(result.stderr)
    print(json.dumps({"file":target.name,"files":len(files),"bytes":target.stat().st_size,"portable_smoke":"passed"}))
