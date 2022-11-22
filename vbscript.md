# VBScript 관련 정리

## 변수

- Dim, Public, Private 문을 사용하여 변수 생성 가능

```VBScript
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

```VBScript

Dim arr(3)
arr(0)="str1"
arr(1)="str2"
arr(2)="str3"

str=arr(0)

<!-- 2차원 배열도 가능 -->
Dim arr(2,4)
```
