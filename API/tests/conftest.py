import pytest


@pytest.fixture
def fake_powershell(tmp_path):
    executable=tmp_path/'fake-powershell'
    executable.write_text('''#!/usr/bin/env python3
import sys,subprocess,time
from pathlib import Path
sys.stdin.readline()
marker=Path(__file__).with_suffix('.running')
child=subprocess.Popen([sys.executable,'-c',"import time;from pathlib import Path;p=Path("+repr(str(marker))+ ");i=0\\nwhile True:\\n p.write_text(str(i));i+=1;time.sleep(.03)"])
time.sleep(60)
''')
    executable.chmod(0o755)
    return executable, executable.with_suffix('.running')

