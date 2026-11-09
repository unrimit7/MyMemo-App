export interface NoteTemplate {
  id: string;
  title: string;
  description: string;
  iconName: 'meeting' | 'daily' | 'project' | 'bug' | 'weekly' | 'idea';
  category: '업무' | '기획' | '개인';
  content: string;
}

export function getNoteTemplates(): NoteTemplate[] {
  const now = new Date();
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
    now.getDate()
  ).padStart(2, '0')} (${days[now.getDay()]})`;

  return [
    {
      id: 'meeting',
      title: '회의록 템플릿',
      description: '참석자, 안건, 결정 사항 및 Action Items',
      iconName: 'meeting',
      category: '업무',
      content: `# [회의록] 안건 제목을 입력하세요

> [!NOTE]
> **회의 개요**
> - **일시**: ${todayStr} 00:00
> - **장소/방식**: 대회의실 / 화상 회의 (Google Meet)
> - **참석자**: 
> - **기록자**: 

---

## 🎯 1. 회의 목적 및 핵심 안건
- 안건 1: 
- 안건 2: 

---

## 💬 2. 논의 내용 & 주요 결정 사항
<details>
<summary>▶ 세부 논의 내용 펼치기</summary>

- **논의 1**: 
- **논의 2**: 
</details>

- **최종 결정 사항**:
  1. 
  2. 

---

## ✅ 3. 실행 과제 (Action Items)
- [ ] 담당자: 해야 할 작업 내용 (~${todayStr}까지)
- [ ] 담당자: 해야 할 작업 내용

---
#회의록 #업무 #ActionItem
`,
    },
    {
      id: 'daily',
      title: '일일 업무 일지',
      description: '오늘의 핵심 목표, 진행 업무 및 특이사항',
      iconName: 'daily',
      category: '업무',
      content: `# [업무일지] ${todayStr}

> [!TIP]
> **오늘의 핵심 목표 (Top 3)**
> 1. 
> 2. 
> 3. 

---

## 📋 1. 진행 업무 (To-Do)
- [x] 오전 루틴 점검 및 메일 회신
- [ ] 
- [ ] 

---

## 🚧 2. 이슈 및 특이사항
- 특이사항이나 협조 요청 사항을 적어주세요.

---

## 🔮 3. 내일 예정 업무
- [ ] 

---
#업무일지 #데일리로그 #할일
`,
    },
    {
      id: 'project',
      title: '프로젝트 기획서',
      description: '프로젝트 개요, 마일스톤 일정 및 R&R',
      iconName: 'project',
      category: '기획',
      content: `# [기획서] 프로젝트 이름

> [!IMPORTANT]
> **프로젝트 핵심 요약**
> - **목표 (KPI)**: 
> - **목표 일정**: ${todayStr} ~ 
> - **총괄 담당자**: 

---

## 📌 1. 배경 및 필요성
- 추진 배경과 기대 효과를 작성하세요.

---

## 🗓️ 2. 추진 마일스톤 (추진 일정)

| 단계 | 주요 작업 내용 | 마감일 | 상태 |
| :--- | :--- | :---: | :---: |
| 기획 | 요구사항 정의 및 화면 설계 | ${todayStr} | 진행중 |
| 개발 | 핵심 기능 구현 및 테스트 | - | 대기 |
| 배포 | 실서버 배포 및 피드백 수렴 | - | 대기 |

---

## 👥 3. 역할 및 책임 (R&R)
- **기획/디자인**: 
- **개발**: 
- **검수**: 

---
#프로젝트 #기획서 #마일스톤
`,
    },
    {
      id: 'bug',
      title: '버그 / 이슈 리포트',
      description: '오류 재현 경로, 원인 분석 및 조치 방안',
      iconName: 'bug',
      category: '업무',
      content: `# [이슈리포트] 오류 및 문제 현상 요약

> [!WARNING]
> **오류 중요도**: 보통 (낮음 / 보통 / 긴급)  
> **발생 일시**: ${todayStr}  
> **발생 환경**: OS: Windows / 모바일 / 브라우저: Chrome

---

## 🔍 1. 현상 및 재현 단계 (Steps to Reproduce)
1. 
2. 
3. 

---

## 📊 2. 예상 동작 vs 실제 동작
- **예상한 정상 동작**: 
- **실제 발생한 오류 동작**: 

---

## 🛠️ 3. 임시 조치 및 해결 방안
- [ ] 원인 분석
- [ ] 패치 적용 및 검증

---
#버그리포트 #이슈 #QA
`,
    },
    {
      id: 'weekly',
      title: '주간 업무 보고',
      description: '금주 주요 실적 및 차주 계획 요약',
      iconName: 'weekly',
      category: '업무',
      content: `# [주간보고] 주간 업무 보고 (${todayStr})

## 🏆 1. 금주 주요 실적 (Done)
- [x] 
- [x] 

---

## 🚀 2. 차주 계획 (Planned)
- [ ] 
- [ ] 

---

## 💡 3. 건의 및 지원 필요 사항
- 건의 사항이나 공유할 소식을 적어주세요.

---
#주간보고 #업무공유
`,
    },
    {
      id: 'idea',
      title: '아이디어 & 기획 메모',
      description: '번뜩이는 아이디어의 배경, 실행 방안 및 메모',
      iconName: 'idea',
      category: '개인',
      content: `# [아이디어] 아이디어 제목

> [!NOTE]
> 한 줄 요약: 어떤 문제를 해결하는 아이디어인가요?

---

## 💡 1. 아이디어 상세 내용
- 구체적인 컨셉과 방식:

---

## 📈 2. 기대 효과 및 장점
- 왜 이 아이디어가 매력적인가?

---

## 👣 3. 첫 번째 실행 단계 (Next Action)
- [ ] 가볍게 시작해볼 수 있는 1단계 작업

---
#아이디어 #기획 #인사이트
`,
    },
  ];
}
