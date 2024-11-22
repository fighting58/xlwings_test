# import xlwings as xw
import numpy as np
from datetime import datetime, timedelta
import xlwings as xw


def convert_angle_to_decimal(angle_str):
    # 입력 문자열에서 각도, 분, 초를 분리
    if not angle_str.endswith('"'):
        angle_str += '"'
        print(angle_str)
    degrees, minutes, seconds = angle_str.split(' ')
    
    # 각도와 분은 정수로 변환
    degrees = int(degrees[:-1])  # 마지막 글자 '˚' 제거
    minutes = int(minutes[:-1])  # 마지막 글자 ''' 제거
    seconds = float(seconds[:-1])  # 마지막 글자 '"' 제거

    # 초를 소수점 4째 자리에서 반올림
    seconds_rounded = np.round(seconds, 4)
    
    # 변환된 값 반환
    return degrees, minutes, seconds_rounded

def convert_decimal_to_angle(degrees, minutes, seconds):
    # 각도, 분, 초를 문자열 형식으로 변환
    angle_str = f"{degrees}˚ {minutes}' {seconds:.4f}\""
    return angle_str


def add_time_to_datetime(datetime_str, hours=0, minutes=0, seconds=0):
    # 문자열을 datetime 객체로 변환
    original_datetime = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")
    
    # timedelta를 사용하여 시간과 분을 더함
    new_datetime = original_datetime + timedelta(hours=hours, minutes=minutes, seconds=seconds)
    
    # 새로운 datetime 객체를 문자열로 변환하여 반환
    return new_datetime.strftime("%Y-%m-%d %H:%M:%S")

def extract_date_from_datetime(datetime_str):
    # 문자열에서 날짜 부분만 추출
    date_str = datetime_str.split(" ")[0]
    return date_str

def format_date_to_korean(datetime_str):
    # 문자열에서 날짜 부분 추출
    date_str = datetime_str.split(" ")[0]
    
    # 날짜를 YYYY, MM, DD로 분리
    year, month, day = date_str.split("-")
    
    # 한국어 형식으로 변환
    formatted_date = f"{year}년 {int(month):02d}월 {int(day):02d}일"
    return formatted_date

def set_outer_border_to_medium(sheet, range_address, thickness=None):
    """주어진 범위의 바깥쪽 테두리를 xlMedium으로 설정하는 함수.

    참고) rng.api.Borders 매핑
        ======================
        7	위쪽 테두리 	xlEdgeTop
        8	아래쪽 테두리	xlEdgeBottom
        9	왼쪽 테두리	    xlEdgeLeft
        10	오른쪽 테두리	xlEdgeRight
        11	내부 수직선	    xlInsideVertical
        12	내부 수평선	    xlInsideHorizontal
        13	대각선 하향선	xlDiagonalDown
        14	대각선 상향선	xlDiagonalUp
    """

    # 범위 가져오기
    rng = sheet.range(range_address)
    
    # 외부 테두리 설정
    for i in range(7, 11):
        rng.api.Borders(i).Weight = thickness  

def set_inner_borders_to_thin(sheet, range_address, thickness=None):
    """주어진 범위의 내부 테두리를 xlThin으로 설정하는 함수."""
    # 범위 가져오기
    rng = sheet.range(range_address)

    for i in range(11, 13):
        rng.api.Borders(i).Weight = thickness  # xlHairline
    
def set_all_borders_to_hairline(sheet, range_address, thickness=None):
    """주어진 범위의 모든 테두리를 xlHairline으로 설정하는 함수."""
    # 범위 가져오기
    rng = sheet.range(range_address)
    
    # 모든 테두리 설정
    for i in range(7, 13):  # xlEdgeTop(7) ~ xlInsideHorizontal(12)
        rng.api.Borders(i).Weight = thickness  # xlHairline

def set_custom_format(sheet, range_address, custom_format):
    """주어진 범위에 사용자 지정 형식을 설정하는 함수."""
    # 범위 가져오기
    rng = sheet.range(range_address)
    
    # 사용자 지정 형식 설정
    rng.number_format = custom_format

def merge_cells(sheet, range_address):
    """지정된 범위의 셀들을 병합하는 함수."""
    # 범위 가져오기
    rng = sheet.range(range_address)
    
    # 셀 병합
    rng.merge()

#===================================



# 엑셀 애플리케이션 시작
app = xw.App(visible=True)

# 데이터 파일 열기
data_wb = app.books.open('data.xlsx')  # data.xlsx 파일 경로 입력
data_ws = data_wb.sheets[0]

# data_ws.range('A1').value = '이름'

# # "Report" 시트가 존재하는지 확인하고 삭제
# if 'Report' in [sheet.name for sheet in data_wb.sheets]:
#     data_wb.sheets['Report'].delete()

# template_sheet = data_wb.sheets['@Report']
# template_sheet.copy(after=data_wb.sheets[-1])  # 마지막 시트 뒤에 복제
# report_sheet = data_wb.sheets[-1]  # 방금 복제한 시트를 가져옴
# report_sheet.name = 'Report'  # 시트 이름 변경


# # # data.xlsx의 데이터 읽기 (2번째 행부터)
# data_range = data_wb.sheets[0].range('A2').expand()  # 첫 번째 시트의 데이터 범위 확장 (헤더 제외)
# data_count = data_range.rows.count  # 데이터의 행 수

# # 템플릿 범위의 값을 A1 셀에 삽입
# current_row = 1  # 현재 행 번호

# # 데이터 복사 및 추가
# for i in range(data_count):
#     # 템플릿 범위를 복사
#     template_range = template_sheet.range('1:17')
#     template_range.copy()
    
#     # A1 셀에 삽입하기 전에 기존 데이터 아래로 밀기
#     report_sheet.range('A1').insert(shift='down')  # A1 셀에 있는 내용을 아래로 밀기
    
#     # 복사한 범위를 A1 셀에 붙여넣기
#     report_sheet.range('A1').paste()

# # data.xlsx의 "B2" 셀의 값을 "Report" 시트의 "B5" 셀에 복사
# value_to_copy = data_wb.sheets[0].range('B2').value  # "B2" 셀의 값 가져오기
# report_sheet.range('B5').value = value_to_copy  # "B5" 셀에 값 복사

set_inner_borders_to_thin(data_ws, "A1:F17", thickness=2)
set_outer_border_to_medium(data_ws, "A1:F17", thickness=3)

# 작업 완료 후 워크북 저장
data_wb.save()
app.quit()  # 엑셀 애플리케이션 종료