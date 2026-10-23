# MinsPDFViewer

WPF(.NET 8) 기반 PDF 뷰어·편집기. PDFium 렌더링, 주석, 전자서명(BouncyCastle), Windows OCR + 검색 가능한 PDF, 인쇄.
기능 현황과 개발 기록은 [doc/개발계획.md](doc/개발계획.md).

## 폴더 구조

```
MinsPDFViewer.sln
src/MinsPDFViewer/          앱 소스 (네임스페이스는 모두 MinsPDFViewer)
  App.xaml, AssemblyInfo.cs
  Views/                    창·대화상자 (MainWindow, PrintWindow, PasswordDialog, 인증서·서명 결과 창)
  Models/                   뷰모델·저장 스냅샷·서명 모델
  Services/                 앱 서비스 (렌더링·로드·저장 조율, 검색, OCR, 인쇄, 열람 기록, 인증서)
  Pdf/                      PDF 구조를 직접 다루는 코드 (PDFium 편집 저장, OCR 텍스트 층, 북마크, 서명·검증, 구문 파서)
  Common/                   변환기·바인딩 도우미·로그
  Assets/                   앱에 포함되는 리소스 (아이콘)
tests/
  MinsPDFViewer.Tests/      xUnit 테스트
  TestFiles/                테스트용 PDF·비교 캡처 (대용량 파일은 git 제외)
doc/                        개발 문서
assets/                     디자인 원본 (로고·아이콘 PNG/ICO, 설치 파일 아이콘)
scripts/                    빌드·배포 스크립트, NSIS 설치 스크립트 (git 추적)
build/                      빌드 산출물 전용 (폴더째 git 제외, 지워도 다시 생김)
  Debug/ Release/           빌드 결과 (프로젝트별 하위 폴더)
  publish/ obj/             게시 결과·중간 파일
lib/                        NuGet에 없는 외부 라이브러리 자리 (현재 비어 있음)
dist/                       설치 파일 출력 (git 제외)
```

## 빌드·실행

```bat
dotnet build MinsPDFViewer.sln
dotnet run --project src\MinsPDFViewer
dotnet test tests\MinsPDFViewer.Tests
```

실행 파일: `build\Debug\MinsPDFViewer\MinsPDFViewer.exe` (Release는 `build\Release\MinsPDFViewer\`)
- 모든 프로젝트의 출력 경로는 루트 `Directory.Build.props`에서 정한다. `src\`·`tests\`에는 소스만 남는다.
- self-contained 빌드라 .NET 런타임이 exe와 같은 폴더에 들어 있다. 테스트 출력 폴더에는 exe를 두지 않는다.

## 배포

```bat
scripts\release.bat      :: publish + 설치 파일 → dist\MinsPDFViewer_Setup_v<버전>.exe
scripts\publish.bat      :: publish만 → build\publish
scripts\makeSetup.bat    :: 설치 파일만 (NSIS 필요: C:\Program Files (x86)\NSIS)
```
