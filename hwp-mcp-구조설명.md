# hwp-mcp 프로젝트 구조 설명

## 개요

hwp-mcp는 한글 워드 프로세서(HWP)를 Cursor AI가 제어할 수 있도록 해주는 Model Context Protocol(MCP) 서버입니다. 이 프로젝트는 한글 문서를 자동으로 생성, 편집, 조작하는 기능을 Cursor의 AI 모델에게 제공합니다.

## 작동 방식

1. **Cursor AI 호출 과정**:
   - Cursor에서 hwp-mcp를 호출하면 `hwp_mcp_stdio_server.py` 스크립트가 실행됩니다.
   - 이 스크립트는 표준 입출력(stdio)을 통해 Cursor AI와 통신하는 FastMCP 서버를 생성합니다.
   - 서버는 HWP를 조작하기 위한 여러 도구(tool)들을 등록합니다.

2. **기능 실행 과정**:
   - Cursor AI가 HWP 관련 기능을 요청하면, 등록된 도구가 호출됩니다.
   - 도구는 `HwpController`를 사용하여 Windows COM 자동화를 통해 HWP 프로그램과 상호작용합니다.
   - `win32com` 라이브러리를 통해 HWP에 명령을 전송합니다.

## 주요 컴포넌트

1. **hwp_mcp_stdio_server.py**:
   - 프로젝트의 진입점(entry point)
   - FastMCP 서버를 초기화하고 도구들을 등록
   - stdio를 통해 Cursor AI와 통신하는 인터페이스 제공

2. **src/tools/hwp_controller.py**:
   - HWP와 상호작용하는 핵심 컨트롤러 클래스
   - win32com을 이용하여 한글 프로그램을 자동화
   - 문서 생성, 열기, 저장, 텍스트 삽입, 테이블 생성 등의 기능 제공

3. **security_module/**:
   - 파일 경로 체크 보안 경고창을 우회하기 위한 모듈 참조
   - 실제 파일은 확인되지 않음

## MCP 도구 목록

1. **문서 관리 도구**:
   - `hwp_create`: 새 문서 생성
   - `hwp_open`: 기존 문서 열기
   - `hwp_save`: 문서 저장

2. **문서 편집 도구**:
   - `hwp_insert_text`: 텍스트 삽입
   - `hwp_insert_table`: 표 생성 및 편집
   - `hwp_create_table_with_data`: 데이터가 포함된 표 생성

## 의존성

1. **Python 패키지 의존성** (requirements.txt):
   - mcp (Cursor AI의 Model Context Protocol 구현체)
   - pywin32==306 (HWP 자동화를 위한 핵심 라이브러리)
   - python-dotenv==1.0.0
   - pydantic==2.5.2
   - jsonschema==4.19.1
   - pytest==7.4.2

## Cursor AI 설정

Cursor AI에서 이 MCP 서버를 사용하기 위해서는 다음과 같은 설정이 필요합니다:

1. MCP 서버 등록:
   - Cursor의 설정에 hwp-mcp 서버 정보 추가
   - 실행 명령어 및 경로 설정

2. 도구 사용 예시:
   ```
   # Cursor AI가 다음과 같은 방식으로 도구를 호출할 수 있습니다
   <function_calls>
   <invoke name="mcp_hwp_hwp_create">
   <parameter name="random_string">dummy
   ```

## 프로젝트 삭제 시 고려사항

이 프로젝트를 삭제하고자 할 경우 다음 사항을 고려해야 합니다:

1. 이 코드는 Cursor AI와 HWP 사이의 브릿지 역할을 합니다.
2. 외부 패키지인 `mcp`에 의존성이 있으며, 이는 별도로 설치되어 있을 가능성이 높습니다.
3. Cursor AI 설정 파일에 이 프로젝트에 대한 참조가 있습니다.

프로젝트를 삭제할 경우 Cursor AI는 더 이상 HWP를 제어할 수 없게 됩니다. 이것이 의도한 바라면, Cursor AI 설정에서 hwp mcp 항목을 제거하는 것도 함께 진행해야 합니다.

## 설정 파일 예시

```json
{
  "mcpServers": {
    "hwp": {
      "command": "python",
      "args": ["D:\\hwp-mcp\\hwp_mcp_stdio_server.py"]
    }
  }
}
```

위 설정에서 "hwp" 항목을 제거하면 Cursor AI가 더 이상 HWP MCP를 호출하지 않습니다. 