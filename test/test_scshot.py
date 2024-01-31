import cv2

# 변수 초기화
drawing = False
ix, iy = -1, -1

# 마우스 이벤트 콜백 함수
def draw_rectangle(event, x, y, flags, param):
    global ix, iy, drawing

    if event == cv2.EVENT_LBUTTONDOWN:
        drawing = True
        ix, iy = x, y
    elif event == cv2.EVENT_MOUSEMOVE:
        if drawing:
            img_copy = img.copy()
            cv2.rectangle(img_copy, (ix, iy), (x, y), (0, 255, 0), 2)
            cv2.imshow('Select Region', img_copy)
    elif event == cv2.EVENT_LBUTTONUP:
        drawing = False
        cv2.rectangle(img, (ix, iy), (x, y), (0, 255, 0), 2)
        cv2.imshow('Select Region', img)
        # 선택한 영역 캡처
        capture = img[iy:y, ix:x]
        cv2.imwrite('C:/Users/user/Desktop/coding/captured_region.png', capture)
        print("영역이 캡처되었습니다.")

# 빈 화면 생성
img = cv2.imread('C:/Users/user/Desktop/coding/captured_region.png')  # 스크린샷 파일 경로 입력
cv2.namedWindow('Select Region')
cv2.setMouseCallback('Select Region', draw_rectangle)

while True:
    cv2.imshow('Select Region', img)
    key = cv2.waitKey(1) & 0xFF
    if key == 27:  # ESC 키를 누르면 종료
        break

cv2.destroyAllWindows()
