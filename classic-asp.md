# Classic ASP 정리

- ASP는 액티브 서버 페이지 약자다
- 확장자는 .asp이다

## asp 파일

- <% 와 %> 사이에는 서버에서 실행한다
- Response.Write()는 HTML출력에 쓰인다
  - VBScript는 대소문자 구별안함
    - response.write()도 가능

## Response.Write() 사용 예

```html
<% Response.Write("글 쓰기")%>
<!-- Response.Write()와 = 같은거 임 -->
<% ="글 쓰기" %> <% Response.Write(
<p stlye="color:#dodgerblue">"글 쓰기"</p>
) %>
```

## 서브 프로시저, 함수 프로시저 선언

- 서브 프로시저 (리턴 없음)

```VB
<!-- 두가지 방법이 있음-->
<!-- 리턴은 없지만 ASP에서 Response로 값을 쓰는 경우다 -->
<!-- 방법 1 -->
<%
Sub fnc(num1, num2)
  Response.write(num1 + num2)
End Sub
%>

<div>Result : <%call fun(2,3) %></div>

<!-- 방법2 -->

<%@ language="javascript" %>
<%
function fnc2(num1,num2)
{
Response.Write(num1 + num2)
}
%>

<div>Result : <% fnc2(3,4) %> </div>
```

- 함수 프로시저 (리턴 있음)

```VB
<%
Function fnc()
  fnc = Date()
  <!-- 함수 명에 값을 넣어준다 -->
end Function
%>
```
