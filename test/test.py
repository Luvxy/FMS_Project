def remove_specific_words(text):
    # 분할된 단어들 중 '통' 또는 '반'을 포함하지 않은 단어들만 선택하여 새 문자열 생성
    result = ' '.join([word for word in text.split() if '통' not in word and '반' not in word])
    return result

# 예제 사용
input_text = "선화로1길 25-14, 403동 1206호 (모현동2가,배산휴먼시아4단지) 51통 1반"
output_text = remove_specific_words(input_text)
print(output_text)