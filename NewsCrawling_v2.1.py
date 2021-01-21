import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
import json  # json 으로 들어오는 결과를 잘 정돈해야지
import re  # 정규표현식으로 태그 떼기위해서 .
import urllib.request
import pandas as pd

homeUI = uic.loadUiType("newGUI.ui 파일의 경로+파일명.ui를 이곳에 넣어주세요.")[0] #newGUI.ui 파일경로와 파일명


class newANDnews(QDialog, homeUI):
    # 클래스 변수로 선언
    keyword_range = ""  # 검색을 계획한 키워드가 1000개고 그 중에 50개만 하고 싶다면? 이런 상황을 대비해서 키워드의 범위를 설정
    max_num = ""  # 최대 몇개의 사이트 정보를 가져올지 결정
    save_name = ""  # 저장할 파일 이름을 입력
    keyword_excel = ""  # 키워드가 담김 엑셀 파일의 위치를 담을 그릇
    folder_path = ""
    ID = ""
    SC = ""
    def __init__(self):  # GUI
        super().__init__()
        self.setupUi(self)

    def openFolder(self):  # 버튼을 클릭해서 폴더의 위치정보를 가져온다.
        newANDnews.folder_path = QFileDialog.getExistingDirectory(self)
        print(newANDnews.folder_path)


    def keywords_open(self):  # 버튼을 클릭해서 검색어 목록이 있는 파일을 가져온다.
        keyword_file = QFileDialog.getOpenFileName(self)
        newANDnews.keyword_excel = keyword_file[0]
        print(newANDnews.keyword_excel)

    def run_news(self):  # 크롤링을 시작한다.
        # 검색 범위 지정해주기
        newANDnews.keyword_range = self.lineEdit1_2.text()
        print("검색 범위 지정해주기", newANDnews.keyword_range)

        # 검색결과의 최대 갯수 정해주기
        newANDnews.max_num = self.lineEdit1_1.text()  # 공란에 입력한 값을 받아온다.
        print("검색결과의 최대 갯수 정해주기", newANDnews.max_num)

        # 저장할 파일명
        newANDnews.save_name = self.lineEdit_save.text()  # 공란에 입력한 값을 받아온다.
        print("저장할 파일명", newANDnews.save_name)


        newANDnews.ID = self.lineEdit_id.text()
        newANDnews.SC = self.lineEdit_secret.text()


        # -----------------------크롤링 영역------------------------------
        # API 계정
        client_id = newANDnews.ID
        client_secret = newANDnews.SC

        # 검색어 엑셀 파일 위치와 크롤링 데이터 저장할 위치
        path_excel_search = newANDnews.keyword_excel
        path_folder = newANDnews.folder_path
        filename = newANDnews.save_name
        path_excel_crawling = path_folder + "/" + filename + ".xlsx"

        idx = 1  # 인덱스를 좀 달자. 0부터 시작
        display = 10  # 한 번에 몇개씩 출력할 것인지, 최대100,  실제로 돌릴때는 최대한 많은 값을 위해 100으로 설정하자
        start = 1  # 시작 위치
        end = int(newANDnews.max_num)  # 최대 1000, 실제로 돌릴때는 최대한 많은 값을 위해 1000으로 설정하자

        # 검색 내용 판다스로 저장하자. 저장하기 위해 데이터 프레임을 columns와 함께 만들어두자.
        web_df = pd.DataFrame(
            columns=("검색어", "기사 제목", "URL 링크"))  # 네이버 API는 현재까지 신문사 명을 제공하지 않고 있어서 부득이 2가지만 먼저 크롤링해본다.

        # 검색어를 리스트를 불러오자.
        searching_list = []  # 검색어를 저장해둘 리스트 초기화
        df = pd.read_excel(path_excel_search, engine='openpyxl')  # 파일 불러오기 # 불러올 파일의 위치를 지정해주세요.
        excel_np = pd.DataFrame.to_numpy(df)  # 엑셀파일을 numpy형식으로 변환

        for_search_limit = int(newANDnews.keyword_range)
        for i in range(0, for_search_limit - 1):
            searching_list.append(excel_np[i][0])
        print(searching_list)

        for searching in searching_list:  # 검색어 리스트에서 검색어를 하나씩 가져온다.
            search_text = urllib.parse.quote("{0}".format(searching))

            # 검색 1개만 할 거 아니니깐
            for start_index in range(start, end, display):  # 예를 들면 검색해서 나온 것 중에 첫 기사(start)부터 10 개씩(display) 300번째 까지(end) 검색
                url = "https://openapi.naver.com/v1/search/news?query=" + search_text \
                      + "&display=" + str(display) \
                      + "&start=" + str(start_index) \
                      + "sort=date"
                # news검색을 위한 url 생성

                request = urllib.request.Request(url)
                request.add_header("X-Naver-Client-Id", client_id)
                request.add_header("X-Naver-Client-Secret", client_secret)
                response = urllib.request.urlopen(request)
                rescode = response.getcode()

                if (rescode == 200): # 200:정상코드. 정상적으로 작동을 한다면 for 구문으로 들어가서 크롤링을 한다.
                    response_body = response.read()
                    response_dict = json.loads(response_body.decode('utf-8'))
                    items = response_dict['items']  # items에 해당하는 결과 모두 가져오기

                    for item_index in range(0, len(items)):
                        remove_tag = re.compile('<.*&;?>')  # 테그 떼어버리기
                        title = re.sub(remove_tag, '', items[item_index]['title'])  # 제목 정보 가져오기
                        link = items[item_index]['link']  # 링크 정보 가져오기
                        web_df.loc[idx] = [searching, title, link]
                        idx = idx + 1
                        with pd.ExcelWriter(path_excel_crawling) as writer:
                            sheetname = "{0}".format(searching)
                            print("시트 이름: ", sheetname)
                            web_df.to_excel(writer, sheet_name=sheetname) #검색어 별로 시트에 저장하고 싶은데...
                            print("검색어 이름과 크롤링 자료 모으기 ->", searching, web_df)
                            print("검색중" , (idx-1) , "/" , int(newANDnews.max_num) * (int(newANDnews.keyword_range)-1) )

                else:
                    print("Error Code:" + rescode)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    homeD = newANDnews()
    homeD.show()
    app.exec_()
