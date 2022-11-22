# Classic ASP 정리

- ASP는 액티브 서버 페이지 약자다
- 확장자는 .asp이다

## asp 파일

- <% 와 %> 사이에는 서버에서 실행한다
- Response.Write()는 HTML출력에 쓰인다
  - VBScript는 대소문자 구별안함
    - response.write()도 가능

### Response.Write() 사용 예

```html
<% Response.Write("글 쓰기")%>
<!-- Response.Write()와 = 같은거 임 -->
<% ="글 쓰기" %> <% Response.Write(
<p stlye="color:#dodgerblue">"글 쓰기"</p>
) %>
```
