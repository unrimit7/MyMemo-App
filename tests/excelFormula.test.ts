import { describe, it, expect } from 'vitest';
import { evaluateFormula, excelSerialFromDate, formatCellValue } from '../src/utils/excelViewer';

describe('excelViewer evaluateFormula 화이트리스트 안전 수식 계산기 검증', () => {
  // 모의 셀 데이터: A1=5, B1=3, C1=1234.5, A4=2026-11-27
  const mockCells: Record<string, any> = {
    A1: 5,
    B1: 3,
    C1: 1234.5,
    A4: new Date(2026, 10, 27),
    D1: '100',
  };

  const resolveCell = (ref: string) => mockCells[ref];

  describe('1. 기존 정규식 치환 제거로 발생하던 7가지 왜곡 수식 검증 (절대 틀린 숫자 반환 금지)', () => {
    it('=MAX(A1,B1) -> 집계 함수 지원: 정확한 5 (기존 53 왜곡 방지)', () => {
      const res = evaluateFormula('=MAX(A1,B1)', resolveCell);
      expect(res.success).toBe(true);
      expect(res.result).toBe(5);
      expect(res.result).not.toBe(53);
    });

    it('=MIN(A1,B1) -> 집계 함수 지원: 정확한 3 (기존 53 왜곡 방지)', () => {
      const res = evaluateFormula('=MIN(A1,B1)', resolveCell);
      expect(res.success).toBe(true);
      expect(res.result).toBe(3);
      expect(res.result).not.toBe(53);
    });

    it('=ROUND(C1,0) -> 화이트리스트 외 함수이므로 success: false (조용한 잘못된 계산 방지)', () => {
      const res = evaluateFormula('=ROUND(C1,0)', resolveCell);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
      expect(res.formulaText).toBe('=ROUND(C1,0)');
    });

    it('=$A$1*2 -> $ 절대참조 정상 파싱하여 정확한 10 반환 (기존 2 왜곡 방지)', () => {
      const res = evaluateFormula('=$A$1*2', resolveCell);
      expect(res.success).toBe(true);
      expect(res.result).toBe(10);
      expect(res.result).not.toBe(2);
    });

    it('=A1^2 -> 거듭제곱(^) 미지원 연산자이므로 success: false (기존 52 왜곡 방지)', () => {
      const res = evaluateFormula('=A1^2', resolveCell);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
      expect(res.result).not.toBe(52);
    });

    it('=A1*10% -> 백분율(%) 미지원 연산자이므로 success: false (기존 50 왜곡 방지)', () => {
      const res = evaluateFormula('=A1*10%', resolveCell);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
      expect(res.result).not.toBe(50);
    });

    it('=A1&B1 -> 문자열 연결(&) 미지원 연산자이므로 success: false (기존 53 왜곡 방지)', () => {
      const res = evaluateFormula('=A1&B1', resolveCell);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
      expect(res.result).not.toBe(53);
    });
  });

  describe('2. 날짜 및 휘발성 함수 (TODAY, NOW, DATE) 계산 안전성 검증', () => {
    it('=A4-TODAY() -> 정확한 잔여 일수(정수) 동적 계산', () => {
      const res = evaluateFormula('=A4-TODAY()', resolveCell);
      expect(res.success).toBe(true);

      const now = new Date();
      const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      const expectedDays = Math.round((mockCells.A4.getTime() - today.getTime()) / (24 * 3600 * 1000));
      expect(res.result).toBe(expectedDays);
    });

    it('=IF(A4-TODAY()<0,"지남",A4-TODAY()) -> IF 조건문 미지원이므로 success: false (왜곡 숫자 반환 금지)', () => {
      const res = evaluateFormula('=IF(A4-TODAY()<0,"지남",A4-TODAY())', resolveCell);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
      expect(res.formulaText).toBe('=IF(A4-TODAY()<0,"지남",A4-TODAY())');
    });

    it('=MAX(0,A4-TODAY()) -> 직접 계산식 인수와 함께 정확히 계산', () => {
      const res = evaluateFormula('=MAX(0,A4-TODAY())', resolveCell);
      const now = new Date();
      const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      const diff = Math.round((mockCells.A4.getTime() - today.getTime()) / (24 * 3600 * 1000));
      expect(res.success).toBe(true);
      expect(res.result).toBe(Math.max(0, diff));
    });

    it('=DATE(2026, 11, 27) - DATE(2026, 11, 20) -> 정확한 일수 차이 7 계산', () => {
      const res = evaluateFormula('=DATE(2026, 11, 27) - DATE(2026, 11, 20)', resolveCell);
      expect(res.success).toBe(true);
      expect(res.result).toBe(7);
    });
  });

  describe('3. 화이트리스트 외 문법 엄격 차단 및 정상 사칙연산 검증', () => {
    it('타 시트 참조(Sheet1!A1) 포함 시 거부', () => {
      const res = evaluateFormula('=Sheet1!A1 + 5', resolveCell);
      expect(res.success).toBe(false);
    });

    it('범위 참조 SUM(A1:A5): 빈칸은 무시하고 숫자·날짜만 합산 (엑셀과 동일)', () => {
      const res = evaluateFormula('=SUM(A1:A5)', resolveCell);
      expect(res.success).toBe(true);
      expect(res.result).toBe(5 + excelSerialFromDate(mockCells.A4));
    });

    it('범위 참조는 함수 인수에서만 허용: =A1:A5 단독 사용 시 거부', () => {
      const res = evaluateFormula('=A1:A5', resolveCell);
      expect(res.success).toBe(false);
    });

    it('숫자로 끝나는 임의 함수명(LOG10(100))이 셀 참조로 오인되지 않고 거부', () => {
      const res = evaluateFormula('=LOG10(100)', resolveCell);
      expect(res.success).toBe(false);
    });

    it('0으로 나누기 발생 시 안전하게 실패 처리 (예외 발생 안 함)', () => {
      const res = evaluateFormula('=A1 / 0', resolveCell);
      expect(res.success).toBe(false);
    });

    it('정상적인 복합 사칙연산과 괄호 우선순위 계산: ((A1 + B1) * 2 - 6) / 2 = 5', () => {
      const res = evaluateFormula('=((A1 + B1) * 2 - 6) / 2', resolveCell);
      expect(res.success).toBe(true);
      expect(res.result).toBe(5); // ((5 + 3) * 2 - 6) / 2 = (16 - 6) / 2 = 5
    });

    it('단항 마이너스(-) 및 $ 참조 변형($A1, A$1) 지원: -$A1 + A$1 = 0', () => {
      const res = evaluateFormula('=-$A1 + A$1', resolveCell);
      expect(res.success).toBe(true);
      expect(res.result).toBe(0);
    });

    it('문자열로 저장된 숫자 셀(D1="100")도 숫자로 변환하여 계산: D1 * 2 = 200', () => {
      const res = evaluateFormula('=D1 * 2', resolveCell);
      expect(res.success).toBe(true);
      expect(res.result).toBe(200);
    });
  });

  describe('4. 집계 함수(SUM/AVERAGE/MIN/MAX/COUNT) — 빈칸·글자 무시, 계산 불가 시 전체 포기', () => {
    // 금액 열 예시: E4~E9 (빈칸, 빈 문자열, 안내 글자 섞임)
    const cells: Record<string, any> = {
      A1: 5,
      E4: 23100,
      E5: undefined,
      E6: 92400,
      E7: '',
      E8: '금액 확인 필요',
      E9: 55000,
      F1: true,
      G1: { formula: 'A1*2' }, // 계산 결과 없는 수식 → 재귀 계산 10
      H1: { formula: 'IF(A1>1,1,0)' }, // 계산할 수 없는 수식(결과 없음)
      H2: { formula: 'IF(A1>1,1,0)', result: 1 }, // 계산 불가지만 저장된 결과 있음
      J1: { error: '#N/A' },
      X1: { formula: 'X1+1' }, // 순환 참조
      K1: { richText: [{ text: '글자' }] },
      L1: 0.1,
      L2: 0.2,
    };
    const rc = (ref: string) => cells[ref];

    it('SUM(E4:E9): 빈칸·빈 문자열·글자는 무시하고 170,500', () => {
      const res = evaluateFormula('=SUM(E4:E9)', rc);
      expect(res.success).toBe(true);
      expect(res.result).toBe(170500);
    });

    it('$ 절대참조 범위와 거꾸로 쓴 범위도 동일: SUM($E$9:E4) = 170,500', () => {
      const res = evaluateFormula('=SUM($E$9:E4)', rc);
      expect(res.success).toBe(true);
      expect(res.result).toBe(170500);
    });

    it('COUNT(E4:E9) = 3 (숫자만 셈), AVERAGE(E4:E9) = 56,833.3333333333', () => {
      expect(evaluateFormula('=COUNT(E4:E9)', rc).result).toBe(3);
      const avg = evaluateFormula('=AVERAGE(E4:E9)', rc);
      expect(avg.success).toBe(true);
      expect(avg.result).toBeCloseTo(170500 / 3, 6);
    });

    it('MIN/MAX(E4:E9) = 23,100 / 92,400, 논리값·서식 있는 글자 참조는 무시', () => {
      expect(evaluateFormula('=MIN(E4:E9)', rc).result).toBe(23100);
      expect(evaluateFormula('=MAX(E4:E9)', rc).result).toBe(92400);
      expect(evaluateFormula('=SUM(F1,K1,A1)', rc).result).toBe(5);
    });

    it('여러 인수와 사칙연산 조합: SUM(E4:E6, A1*2, 10) + 1 = 115,521', () => {
      const res = evaluateFormula('=SUM(E4:E6, A1*2, 10) + 1', rc);
      expect(res.success).toBe(true);
      expect(res.result).toBe(23100 + 92400 + 10 + 10 + 1);
    });

    it('계산 결과 없는 수식 셀은 재귀 계산: SUM(G1,A1) = 15', () => {
      const res = evaluateFormula('=SUM(G1,A1)', rc);
      expect(res.success).toBe(true);
      expect(res.result).toBe(15);
    });

    it('계산할 수 없고 결과도 없는 수식이 범위에 있으면 전체 계산 포기 (틀린 합계 금지)', () => {
      const res = evaluateFormula('=SUM(H1,A1)', rc);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
    });

    it('계산은 못 하지만 저장된 결과가 있는 수식은 그 값을 사용: SUM(H2,A1) = 6', () => {
      const res = evaluateFormula('=SUM(H2,A1)', rc);
      expect(res.success).toBe(true);
      expect(res.result).toBe(6);
    });

    it('오류 값(#N/A)이 범위에 있으면 계산 포기', () => {
      expect(evaluateFormula('=SUM(J1,A1)', rc).success).toBe(false);
    });

    it('순환 참조는 깊이 제한으로 안전하게 실패', () => {
      expect(evaluateFormula('=X1', rc).success).toBe(false);
      expect(evaluateFormula('=SUM(X1)', rc).success).toBe(false);
    });

    it('숫자가 하나도 없는 AVERAGE는 #DIV/0! → 계산 포기, SUM은 0', () => {
      expect(evaluateFormula('=AVERAGE(E5:E5)', rc).success).toBe(false);
      // E7(빈 문자열), E8(안내 글자)만 있는 범위 → 0
      expect(evaluateFormula('=SUM(E7:E8)', rc).result).toBe(0);
    });

    it('비슷한 이름의 미지원 함수(SUMIF, COUNTA, MAXA)와 타 시트 범위는 거부', () => {
      expect(evaluateFormula('=SUMIF(E4:E9,">0")', rc).success).toBe(false);
      expect(evaluateFormula('=COUNTA(E4:E9)', rc).success).toBe(false);
      expect(evaluateFormula('=MAXA(E4:E9)', rc).success).toBe(false);
      expect(evaluateFormula('=SUM(Sheet1!E4:E9)', rc).success).toBe(false);
      expect(evaluateFormula('=SUM()', rc).success).toBe(false);
    });

    it('함수 밖 쉼표는 거부: =A1,E4', () => {
      expect(evaluateFormula('=A1,E4', rc).success).toBe(false);
    });

    it('부동소수점 오차 정리: SUM(L1:L2) = 0.3', () => {
      expect(evaluateFormula('=SUM(L1:L2)', rc).result).toBe(0.3);
    });
  });
});
