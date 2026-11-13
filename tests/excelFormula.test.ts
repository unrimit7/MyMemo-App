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
    it('=MAX(A1,B1) -> 화이트리스트 외 함수이므로 success: false (기존 53 왜곡 방지)', () => {
      const res = evaluateFormula('=MAX(A1,B1)', resolveCell);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
      expect(res.formulaText).toBe('=MAX(A1,B1)');
      // 절대로 53이나 왜곡된 숫자가 나와서는 안 됨
      expect(res.result).not.toBe(53);
    });

    it('=MIN(A1,B1) -> 화이트리스트 외 함수이므로 success: false (기존 53 왜곡 방지)', () => {
      const res = evaluateFormula('=MIN(A1,B1)', resolveCell);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
      expect(res.formulaText).toBe('=MIN(A1,B1)');
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

    it('=MAX(0,A4-TODAY()) -> MAX 함수 미지원이므로 success: false (왜곡 숫자 반환 금지)', () => {
      const res = evaluateFormula('=MAX(0,A4-TODAY())', resolveCell);
      expect(res.success).toBe(false);
      expect(res.result).toBeUndefined();
      expect(res.formulaText).toBe('=MAX(0,A4-TODAY())');
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

    it('범위 참조(A1:B2) 포함 시 거부', () => {
      const res = evaluateFormula('=SUM(A1:A5)', resolveCell);
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
});
