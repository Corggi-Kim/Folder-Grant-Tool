"""Capture the real Qt app with offline sample records matching the approved mockup.

Run from repository root:
QT_QPA_PLATFORM=offscreen Client/.venv/bin/python Client/tests/render_ui_preview.py
No API, BUS, AD or ACL operations are performed.
"""
import datetime
import json
from pathlib import Path
from test_ui_flow import UiFlowTests


def render():
    fixture = UiFlowTests()
    fixture.setUpClass()
    fixture.setUp()
    try:
        w,m,app = fixture.window,fixture.m,fixture.app
        w.api_config['base_url'] = 'http://192.168.1.10:8000'
        # Notification/count badges are disabled in this client.
        w.session.is_ready = lambda:True
        w._set_api_connection_state(True,dict(executor_mode='powershell'))
        started = (datetime.datetime.now(datetime.timezone.utc)-datetime.timedelta(seconds=522)).isoformat()
        steps = [dict(type='ensure_group',status='succeeded'),
                 dict(type='add_group_member',status='succeeded'),dict(type='grant_acl',status='running')]
        records = [
            ('진행','권한부여','10234','26-012','Study','STAT','Manager','API 실행중',
             dict(job_id='sample-1',status='running',executor_mode='powershell',started_at=started,steps=steps)),
            ('종료','권한해제','10381','21-040','','','','대기',{}),
            ('진행','권한부여','10142','25-074','Isolated','STAT_IDMC','','API 성공',
             dict(job_id='sample-3',status='succeeded',executor_mode='powershell')),
            ('종료','권한해제','10567','20-018','','','','서버 시간 초과 (BUS 완료 보류)',
             dict(job_id='sample-4',status='failed',executor_mode='powershell',error_code='POWERSHELL_TIMEOUT')),
            ('진행','권한부여','10203','26-006','Study','DM','','BUS 완료',
             dict(job_id='sample-5',status='succeeded',executor_mode='powershell')),
            ('진행','권한해제','10451','25-077','Study','STAT','','DryRun',dict(status='preview',valid=True)),
        ]
        for index,(kind,request,user,project,lv2,lv3,role,status,result) in enumerate(records):
            w.add_table_row(request,user,project,lv2,lv3,'',role,kind)
            w._set_row_metadata(index,dict(source='bus',request_id=f'BUS-sample-{index}',result=result,bus_done=status=='BUS 완료'))
            w.table.item(index,w.COL_STATUS).setText(status)
            w._get_checkbox(index).setChecked(index<2)
        w._mark_list_loaded('access',w.table.rowCount())
        w.log.setPlainText('10:21:04  요청 검증 완료\n10:21:05  그룹 권한 적용 완료\n10:21:06  폴더 권한 적용 중…')
        w.table.setCurrentCell(0,w.COL_PROJ)
        w.btn_stop.setEnabled(True)
        w.resize(1060,900)
        w.show()
        output = Path(__file__).parents[2]/'docs/screenshots'
        output.mkdir(parents=True,exist_ok=True)
        for mode in ('light','dark'):
            w.apply_theme(mode)
            for _ in range(4):app.processEvents()
            w.grab().save(str(output/f'client-{mode}-v1.1.0.png'))
        if '--projects' in __import__('sys').argv:
            w.ensure_bus_session_async = lambda purpose,callback:None
            w.open_new_viewer()
            w._newdlg.set_data(['프로젝트코드','프로젝트명','신청자','요청일'],[['26-012','표준 프로젝트','홍길동','2026-10-22'],['26-019','신규 연구','김연구','2026-10-22']])
            w._mark_list_loaded('project',w._newdlg.tbl.rowCount())
            app.processEvents()
            w.grab().save(str(output/'client-projects-v1.1.0.png'))
        if '--error' in __import__('sys').argv:
            w.tabs.setCurrentIndex(0); w.apply_theme('light')
            error = '\\\\LSK_S010\\Study Folder\\19087\\Study\\MW: The system cannot find the file specified.\nicacls grant failed: 2'
            w.table.item(0,w.COL_PROJ).setText('1987')
            w.table.item(0,w.COL_LV3).setText('MW')
            w._set_row_metadata(0,dict(source='bus',result=dict(status='partially_succeeded',error_code='POWERSHELL_FAILED',error_message=error,
                steps=[dict(type='add_ad_group_member',status='succeeded'),dict(type='grant_folder_acl',status='failed')])))
            w.table.item(0,w.COL_STATUS).setText('BUS 완료 보류: POWERSHELL_FAILED')
            w.table.setCurrentCell(0,w.COL_PROJ); w._refresh_detail()
            for _ in range(4):app.processEvents()
            w.grab().save(str(output/'client-error-detail-v1.1.0.png'))
        if '--busy' in __import__('sys').argv:
            w.tabs.setCurrentIndex(0); w._set_running_ui(True)
            for mode in ('light','dark'):
                w.apply_theme(mode)
                for _ in range(4):app.processEvents()
                w.grab().save(str(output/f'client-busy-{mode}-v1.1.0.png'))
            w._set_running_ui(False)
        if '--login' in __import__('sys').argv:
            login = m.BusLoginDialog('',w)
            login.show()
            for _ in range(4):app.processEvents()
            login.grab().save(str(output/'client-startup-login-v1.1.0.png'))
            login.hide()
        if '--startup' in __import__('sys').argv:
            from PyQt5.QtTest import QTest
            w.hide()
            dialog = m.StartupLoadingDialog(w)
            dialog.show()
            frames = Path('/tmp/fgt-startup-frames'); frames.mkdir(exist_ok=True)
            frame = 0
            for stage in range(4):
                dialog._stage(stage)
                for tick in range(12):
                    QTest.qWait(50)
                    dialog.grab().save(str(frames/f'{frame:03}.png'))
                    frame += 1
                if stage==1:
                    dialog.grab().save(str(output/'client-startup-v1.1.0.png'))
            dialog.deadline.stop(); dialog.hide()
        print(json.dumps(dict(size=[w.width(),w.height()],table=[w.table.width(),w.table.height()],
                              detail_width=w.detail_panel.width(),output=str(output)),ensure_ascii=False))
    finally:
        fixture.tearDown()


if __name__=='__main__':
    render()
