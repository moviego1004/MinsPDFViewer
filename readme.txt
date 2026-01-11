#테스트 케이스 만들기 
# 1. 테스트 프로젝트 생성 (xUnit)
dotnet new xunit -n MinsPDFViewer.Tests

# 2. 솔루션에 테스트 프로젝트 추가
dotnet sln add MinsPDFViewer.Tests/MinsPDFViewer.Tests.csproj

# 3. 테스트 프로젝트가 메인 프로젝트(MinsPDFViewer)를 참조하도록 설정
dotnet add MinsPDFViewer.Tests/MinsPDFViewer.Tests.csproj reference MinsPDFViewer/MinsPDFViewer.csproj

# 4. 테스트 프로젝트에 필요한 라이브러리 설치 (메인과 동일환경 구성)
dotnet add MinsPDFViewer.Tests/MinsPDFViewer.Tests.csproj package Docnet.Core
dotnet add MinsPDFViewer.Tests/MinsPDFViewer.Tests.csproj package PdfSharp