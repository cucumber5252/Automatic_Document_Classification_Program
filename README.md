# 공문 분류 자동화 프로그램

## 상세 설명

- 2023년 개발을 막 배우기 시작한 당시 사회복무요원 업무 중 각 파일 별로 스캔을 하고 엑셀에 제목을 달아주던 업무의 효율성을 개선하기 위해 개발했습니다.
- 엑셀에 적절한 정보를 기입되었고 스캔 과정에서 누락된 페이지가 없다는 가정 하에 한번에 스캔된 파일을 공문 별로 나눠주고, 각 공문 파일의 제목이 자동으로 변경된 후 센터 서버의 적절한 위치에 들어가도록 해줍니다.

## 개발 기간

2023.04

## 주요 스택

python, sellenium

## 사용법

1. 엑셀 공문에 처리하고자 하는 공문의 정보를 기입합니다.
2. 공문 별로 스캔하지 않고 "한번에" 처리하고자 하는 공문을 하나의 파일로 스캔합니다.
3. 처리한 하나의 파일을 "파일 분할" 폴더에 넣고 "new_split_file.py"를 엽니다.
4. "new_split_file.py" 파일 안의 ##### 부분을 참고하여 파일에 적절한 정보를 기입한 후 오른쪽 위의 재생 버튼을 누릅니다.
5. 잠시 기다리면 자동으로 ilovepdf 웹사이트가 작동되고 해당 웹사이트에 적절한 정보가 담기고 분할된 pdf 파일이 압축 해제되어 "스캔 파일 제목 변경" 폴더에 담기는 걸 확인할 수 있습니다.
6. 3의 과정처럼 이번엔 "change_file_name.py"를 열고 ##### 부분을 참고하여 파일에 적절한 정보를 기입하면 서버의 적절한 위치에 제목이 변경된 파일이 저장된 걸 확인하실 수 있습니다.

## 주의사항

- 해당 프로그램은 광주시 수어통역센터 공문 분류 시스템 자동화를 위해 개발되었습니다.
- 수신 문서와 발신 문서의 경우 페이지 범위를 입력할 때 "비고"란에 직접 (수신의 경우 I열, 발신의 경우 J열) 입력하지 마시고, 비고 뒤에 숨겨진 "비고 기재" 란에 각 공문 시작 페이지의 첫 페이지 번호를 입력하신 후 비고와 페이지 수는 상위 행에 지정된 함수가 자동 적용(해당 셀 오른쪽 아래에 커서를 대면 +모양이 나오는데 이걸 그대로 아래 행으로 드래그하시면 됩니다) 되어 "비고 기재" 란에 입력된 내용을 바탕으로 "비고"란과 "페이지 수" 란이 자동 입력될 수 있도록 해주셔야 프로그램이 정상 작동합니다.
- 연도가 바뀔 경우 서버에서 엑셀 파일이 존재하는 위치가 변경 되기에 엑셀 파일을 불러오는 경로를 수정해야 합니다.
- 이 프로그램은 VS code, 파이썬 등이 깔려있는 사회복무요원 컴퓨터에서만 구동 가능하기에 이 컴퓨터 이외의 센터 컴퓨터에서는 정상 작동하지 않습니다.
- 해당 프로그램은 2023년경 근무하던 사회복무요원이 업무 자동화를 위해 만든 파일이기에 다른 직원 분들께 여쭤보셔도 아예 모르실 겁니다. 만약 코딩을 해보신 적 없고, 계속 에러가 발생하며, 챗gpt에서도 해답을 얻지 못했다면 그냥 수작업으로 처리하시는 게 빠르실 겁니다!
