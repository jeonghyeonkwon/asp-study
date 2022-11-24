# VBScript 관련 정리

## 변수

- Dim, Public, Private 문을 사용하여 변수 생성 가능

```VB
Dim x
Public y
Private z
x="hello world"

<!-- 변수 이름을 틀리게 써서 새변수 생성이 되는 것을 막기 위해 스크립트 맨 위에 Option Explicit문을 붙여준다 -->

Option Explicit
Dim str;
str"hello world!"
```

## 배열 선언

```VB

Dim arr(3)
arr(0)="str1"
arr(1)="str2"
arr(2)="str3"

str=arr(0)

<!-- 2차원 배열도 가능 -->
Dim arr(2,4)
```

## 조건문

- if문

```VB
<!-- IF ~하다면 THEN ~해라 ElseIF ~그렇지 않고 ~하다면 THEN ~해라 ELSE 아무것도 아니라면 이거해라 End IF  -->
IF i = 1 THEN Response.write("1")
ELSEIF i = 2 THEN Response.write("2")
ELSEIF i = 3 THEN Response.write("3")
ELSE Response.write("4")
End IF
```

- Select ~ Case문(Switch 문이랑 비슷)

```VB
d = 3
Select Case d
    Case 1
        Response.write("1")
    Case 2
        Response.write("2")
    Case 3
        Response.write("3")
    Case Else
        Response.write("4")
End Select
```
