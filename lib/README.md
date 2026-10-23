# lib

NuGet으로 받을 수 없는 외부 라이브러리(DLL 등)를 직접 넣는 곳입니다. 지금은 비어 있습니다.

현재 쓰는 라이브러리는 모두 NuGet 패키지라 빌드할 때 자동으로 내려받습니다 (`src/MinsPDFViewer/MinsPDFViewer.csproj`).

| 패키지 | 용도 |
|---|---|
| PdfiumViewer 2.13.0 | PDF 렌더링·인쇄 (.NET 래퍼) |
| PdfiumViewer.Native.x86_64.v8-xfa 2018.4.8.256 | pdfium.dll (x64 네이티브) |
| Portable.BouncyCastle 1.9.0 | 전자서명(CMS)·타임스탬프 |
| System.Drawing.Common 8.0.0 | 비트맵·인쇄(System.Drawing.Printing) |
| System.Text.Encoding.CodePages 6.0.0 | 한글 코드페이지 |

여기에 DLL을 넣으면 `.gitignore`의 `*.dll` 규칙 예외(`!lib/**/*.dll`)로 git에 포함됩니다.
csproj에서는 `<Reference Include="..\..\lib\파일.dll" />`처럼 참조하세요.
