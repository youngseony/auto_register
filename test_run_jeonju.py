# -*- coding: utf-8 -*-
"""전주 교육과정 시험 모드 테스트 스크립트"""
import logging
import os
import sys
from pathlib import Path

# UTF-8 출력 설정
sys.stdout.reconfigure(encoding='utf-8')

import config
config.DRY_RUN = True  # 확실하게 시험 모드 강제

import tasks.edu_plan as edu_plan
from common.console import setup_logging

def test_main():
    setup_logging()
    csv_path = Path("샘플/전주_교육과정_테스트.csv")
    print(f"\n[시험 모드 테스트] 파일 로드: {csv_path}")
    loaded = edu_plan.load([csv_path])
    # 보조강사 실명 테스트용: 환경변수 TEST_ASSISTANT 로만 받고 파일/로그에는 남기지 않는다.
    name = os.environ.get("TEST_ASSISTANT", "").strip()
    if name:
        masked = name[0] + "*" * (len(name) - 1)

        class _Mask(logging.Filter):
            def filter(self, record):
                record.msg = str(record.getMessage()).replace(name, masked)
                record.args = ()
                return True

        for h in logging.getLogger().handlers:
            h.addFilter(_Mask())
        for r in loaded.rows:
            r[config.COL_ASSISTANT] = name
    edu_plan.show_summary(loaded)
    print("\n브라우저를 실행하고 로그인 대기 상태로 진입합니다...")
    edu_plan.run(loaded)

if __name__ == "__main__":
    test_main()
