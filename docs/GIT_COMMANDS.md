# 깃 & 깃허브(Git & GitHub) 핵심 명령어 가이드

이 문서는 Git과 GitHub를 사용할 때 자주 쓰이는 핵심 명령어와 실무 작업 흐름을 정리한 치트시트(Cheat Sheet)입니다.

---

## 1. 기본 일상 작업 흐름 (Daily Workflow)

코드 수정 후 GitHub에 반영할 때 매일 반복하는 기본 3단계 흐름입니다.

```cmd
# 1. 변경된 파일 목록 및 상태 확인
git status

# 2. 구체적인 코드 변경 내용(diff) 확인
git diff

# 3. 커밋할 파일들을 스테이지(장바구니)에 올리기
git add .              # 변경된 모든 파일 추가
git add <파일명>        # 특정 파일만 선택해서 추가

# 4. 로컬 저장소에 커밋(기록) 생성
git commit -m "작업 내용 요약 설명"

# 5. GitHub 원격 저장소로 업로드 (main 브랜치 기준)
git push origin main
```

---

## 2. GitHub 원격 저장소와 동기화 (Sync)

다른 사람의 작업이나 다른 기기에서 올린 최신 코드를 내 컴퓨터로 가져올 때 사용합니다.

| 명령어 | 설명 | 예시 |
| :--- | :--- | :--- |
| **`git clone <URL>`** | GitHub 저장소를 내 컴퓨터에 처음 복제 | `git clone https://github.com/unrimit7/MyMemo-App.git` |
| **`git pull origin <브랜치>`** | GitHub의 최신 변경 내용을 내려받아 내 코드에 병합 | `git pull origin main` |
| **`git fetch origin`** | GitHub의 변경 내역만 확인하고, 내 코드에 바로 합치지는 않음 | `git fetch origin` |
| **`git remote -v`** | 현재 연결된 원격 GitHub 저장소 주소 확인 | `git remote -v` |

---

## 3. 브랜치(Branch) 관리 (기능 분기 및 협업)

새로운 기능을 개발하거나 실험할 때 기존 `main` 코드를 건드리지 않고 별도 가지를 만들어 작업합니다.

| 명령어 | 설명 | 예시 |
| :--- | :--- | :--- |
| **`git branch`** | 로컬 브랜치 목록 조회 (현재 위치 확인) | `git branch` |
| **`git branch -a`** | 원격(GitHub) 브랜치까지 모두 조회 | `git branch -a` |
| **`git switch -c <이름>`** | 새 브랜치를 생성하고 즉시 이동 | `git switch -c feat/dark-mode` |
| **`git switch <이름>`** | 다른 브랜치로 전환 | `git switch main` |
| **`git merge <이름>`** | 다른 브랜치 작업 내용을 현재 브랜치로 합치기 | `git merge feat/dark-mode` |
| **`git branch -d <이름>`** | 작업이 끝나 병합된 브랜치 삭제 | `git branch -d feat/dark-mode` |

---

## 4. 작업 취소 및 되돌리기 (Undo & Restore)

실수했거나 변경 사항을 원복하고 싶을 때 사용하는 안전장치입니다.

| 상황 | 명령어 | 설명 |
| :--- | :--- | :--- |
| **파일 수정을 원래대로 취소** | `git restore <파일명>` | 마지막 커밋 시점의 원본 상태로 파일 복구 |
| **모든 파일 수정 일괄 취소** | `git restore .` | 현재 작업 트리의 모든 수정을 취소 |
| **`git add` 취소 (스테이지 내리기)** | `git restore --staged <파일명>` | 커밋 장바구니에서만 내리고 파일 내용은 유지 |
| **최근 커밋 1개 취소** | `git reset HEAD~1` | 커밋만 취소하고 작성한 코드는 그대로 유지 |
| **작업 임시 보관 (치워두기)** | `git stash` | 작업 중인 미완성 코드를 임시 저장소에 숨김 |
| **임시 보관한 작업 다시 꺼내기** | `git stash pop` | 가장 최근에 `stash`한 작업을 복원 |

---

## 5. 커밋 기록 및 로그 확인 (Log)

| 명령어 | 설명 |
| :--- | :--- |
| **`git log --oneline`** | 커밋 기록을 한 줄씩 축약해서 간결하게 보기 |
| **`git log -n 5`** | 최근 5개의 커밋 상세 내역(작성자, 날짜, 전체 메시지) 보기 |
| **`git log --graph --oneline`** | 브랜치 분기와 병합 흐름을 텍스트 그래프로 시각화 |

---

## 6. Windows CMD 환경 추천 커밋 체크리스트

MyMemo 프로젝트 작업 시 권장하는 표준 터미널 절차입니다.

```cmd
:: 1. 작업 폴더 이동
cd /d D:\Tool\MyMemo

:: 2. 테스트 및 프로덕션 빌드 통과 확인
npm test
npm run build

:: 3. 상태 점검 (개인정보, 실사용자 데이터 포함 여부 확인)
git status
git diff

:: 4. 스테이징 및 커밋
git add .
git commit -m "feat/fix: 커밋 메시지 작성"

:: 5. GitHub 푸시 (GitHub Actions 자동 배포 트리거)
git push origin main
```
