[ 실행 방법 ]
0. 로컬 디렉토리로 '음성 및 언어 엑셀 추출 프로그램' 전 폴더 복사 ( NAS 직접 실행시 속도 너무 느림 )
1. ResourceFileChecker.exe 실행
2. 우분투 홈 디렉토리를 자신의 우분투 홈 디렉토리로 설정
3. 엑셀 결과 저장 폴더를 결과 엑셀 파일을 받고싶은 디렉토리로 설정
4. 두 디렉토리가 모두 유효한 디렉토리라면 하단의 Run 버튼 활성화 -> 실행하면 됨


[ 옵션 ]
- 실행할 프로그램 :
	1. TTS Analysis : TTS 음성파일을 '프로그램 실행시 기준값' 기준으로 파일 유무 전 프로젝트 분석 및 엑셀로 추출
	2. Locale XML Analysis : locale_xx.xml 파일들을 '프로그램 실행시 기준값' 기준으로 해당 ui 언어 존재 유무 체크 및 엑셀로 추출

- 프로그램 실행시 기준값 : 
	TTS 나 locale 파일을 비교분석할 때 기준이 되는 프로젝트/언어 선택하는 옵션
