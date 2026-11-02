from __future__ import annotations
import os
for key in ('OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS'):
    os.environ[key]='1'
import argparse
import json
import multiprocessing
import sys
from pathlib import Path

def main():
    for stream in (sys.stdout,sys.stderr):
        if stream is not None and hasattr(stream,'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser=argparse.ArgumentParser(description='HYBE PDF 자동 입력 v1.3.1 — 원문 확정값만 입력, 미확정은 빈칸')
    parser.add_argument('--version',action='version',version='HybeFiller 1.3.1')
    parser.add_argument('--pdf',nargs='+',type=Path)
    parser.add_argument('--template',type=Path)
    parser.add_argument('--output',type=Path)
    parser.add_argument('--reference',type=Path)
    parser.add_argument('--reviewed',type=Path,help='PDF SHA256과 원문 페이지를 포함한 직접 검토확정 JSON')
    parser.add_argument('--cache-dir',type=Path,help='진단용 원문 추출 캐시 폴더 (결과 확정값 저장소가 아님)')
    parser.add_argument('--dry-run',action='store_true')
    parser.add_argument('--allow-duplicates',action='store_true')
    parser.add_argument('--json',action='store_true')
    args=parser.parse_args()
    if not args.pdf:
        from app.gui import launch_gui
        launch_gui()
        return 0
    if not args.template or not args.output:
        parser.error('--pdf 사용 시 --template와 --output이 필요합니다.')
    from app.service import run_batch,serialize
    try:
        result=run_batch(args.pdf,args.template,args.output,reference_path=args.reference,reviewed_path=args.reviewed,
                         dry_run=args.dry_run,allow_duplicates=args.allow_duplicates,cache_dir=args.cache_dir,
                         progress=None if args.json else lambda s:print(s,flush=True))
        print(json.dumps(result,ensure_ascii=False,default=serialize,indent=2))
        return 2 if result['counts']['error'] or result['cancelled'] else 0
    except Exception as exc:
        if args.json:
            print(json.dumps({'error':str(exc)},ensure_ascii=False))
        else:
            print(str(exc),file=sys.stderr)
        return 1

if __name__=='__main__':
    multiprocessing.freeze_support()
    sys.exit(main())
