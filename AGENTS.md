# Repo AGENTS

> Scope: repository-local rules
> Global policy: `~/.codex/AGENTS.md`
> Human master guide: `~/agent-coding/agent-system/A2-workspace-memory/Guide.md`
> Structure reference: `~/agent-coding/agent-system/A2-workspace-memory/Structure.md`

## Repository
- current repo: `kofia-am-profile`
- workspace path: `~/agent-coding/agent-projects/A4-worker-repos/kofia-am-profile`
- standardized project name: `kofia_am_profile`

## Objective
KOFIA 자산운용사 종합 프로파일: FreeSIS 설정규모 데이터와 KOFIA 회원사 정보를 조인하여 통합 엑셀 + 대시보드 + 검색 기능 제공

## Usage rules
1. Follow the global rules in `~/.codex/AGENTS.md`.
2. Read A2 workspace-memory docs when project-wide context is needed.
3. Keep this file focused on repository-specific rules only.
4. For new repos, treat `~/agent-coding/agent-system/A1-system-governance/docs/TARGET_OS/` as the target-OS governance baseline.

## Repository-specific rules
- 모든 엑셀 수치 컬럼은 RAW 시트를 Excel 함수로 참조 (하드코딩 금지)
- RAW 데이터는 절대 수정하지 않음 (원본 보존)
- FreeSIS API 호출 시 세션 부트스트랩 + 지연 시간 준수
- 조인 키는 회사명 (strip 정규화)

## Key files
- `README.md`
- `ENTRY.md`
- `AGENTS.md`
- `STATUS.md`
- `src/kofia_am_profile/cli.py`
