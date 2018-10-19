# ApachePoi paractice

## 설정
1. Response의 ContentType을 'test/plain'으로 설정한다.
2. Response의 CharacterEncoding을 'UTF-8'로 설정한다.

## 자바소스
1. Workbook 객체 생성
추출 파일 형태에 맞는 Workbook 클래스 객체를 생성 후  출력 형태 클래스로 형변환해서 담는다.

예) Woork wb = new HSSFWorkbook(); // HSSF는 엑셀 전용 (97이후 버전부터 전부 가능)

2. Sheet 객체 생성
데이터를 담는 Sheet 객체 생성

예) HSSFSheet sheet = (HSSFSheet) wb.createSheet();

3. 

셀 병합 = CellRangeAddress(시작 행, 끝나는 행, 시작 열, 끝나는 열)
// 행은 1부터 시작 열은 0부터 시작



