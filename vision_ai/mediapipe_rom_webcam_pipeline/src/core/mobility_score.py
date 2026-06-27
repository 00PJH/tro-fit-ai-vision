"""
mobility_score.py — Tro-Fit 관절 가동성 등급 평가 엔진
========================================================

[적용 범위]
  normal_rom.json에 정의된 7개 동작만 지원:
    어깨: shoulder_flexion / shoulder_extension / shoulder_abduction
    팔꿈치: elbow_flexion / elbow_extension
    무릎: knee_flexion / knee_extension

  고관절(hip)·발목(ankle)은 MediaPipe 벡터 정의와 해부학적 임상값이
  달라 신뢰성 있는 등급화가 불가능하므로 1차 개발 범위에서 제외합니다.

[논문 근거]
  어깨 굴곡/외전 : Namdari S et al. J Shoulder Elbow Surg. 2012;21(9):1177-83.
  어깨 신전      : AMA 기준값 기반 (ADL 전용 확정 논문 없음, 비율 등분)
  팔꿈치 굴곡/신전: Morrey BF et al. J Bone Joint Surg Am. 1981;63(6):872-7.
  무릎 굴곡      : Rowe PJ et al. Gait Posture. 2000;12(2):143-55.
  무릎 신전      : 임상 관행 기반 (완전 신전 불가 시 보행 장애)
  좌우 비대칭    : 임상 관행 기반 (15° 경계값)

[등급 방향 정의]
  higher_is_better (굴곡 계열):
    측정 ROM이 클수록 좋음. min_deg 이상이면 상위 등급.
    예: 무릎 굴곡 120° 이상 → NORMAL

  lower_is_better (신전 계열):
    측정값(구축각)이 작을수록 좋음. max_deg 이하이면 상위 등급.
    예: 무릎 신전 10° 이하 잔류 → NORMAL (완전 신전에 가까움)

[사용법]
  from core.mobility_score import calculate_mobility_score

  result = calculate_mobility_score({
      "joint":     "knee_flexion",
      "left_rom":  118.5,
      "right_rom": 105.2,
  })

[출력 스키마]
  {
    "joint":           str,
    "left_rom":        float | None,
    "right_rom":       float | None,
    "left_grade":      "NORMAL" | "WARNING" | "CRITICAL",
    "right_grade":     "NORMAL" | "WARNING" | "CRITICAL",
    "overall_grade":   "NORMAL" | "WARNING" | "CRITICAL",
    "asymmetry_deg":   float | None,
    "asymmetry_alert": bool,
    "clinical_meaning": str,
    "recommendation":  str,
    "adl_ok":          list[str],
    "adl_limited":     list[str],
    "scoring_basis":   str,
  }
"""

from __future__ import annotations
from typing import Literal

# ─────────────────────────────────────────────────────────────────────────────
# 타입 정의
# ─────────────────────────────────────────────────────────────────────────────

Grade = Literal["NORMAL", "WARNING", "CRITICAL"]


# ─────────────────────────────────────────────────────────────────────────────
# 등급 기준 테이블
# ─────────────────────────────────────────────────────────────────────────────
# 형식:
#   "higher_is_better" 동작 (굴곡 계열):
#     (NORMAL_min, WARNING_min)
#     ROM >= NORMAL_min → NORMAL
#     ROM >= WARNING_min → WARNING
#     ROM < WARNING_min → CRITICAL
#
#   "lower_is_better" 동작 (신전 계열):
#     (NORMAL_max, WARNING_max)  ← 구축각(잔류 굴곡)
#     구축각 <= NORMAL_max  → NORMAL
#     구축각 <= WARNING_max → WARNING
#     구축각 > WARNING_max  → CRITICAL

_JOINTS: dict[str, dict] = {

    # ── 어깨 굴곡 (Namdari 2012) ──────────────────────────────────────────
    "shoulder_flexion": {
        "direction": "higher_is_better",
        "thresholds": (120.0, 90.0),   # (NORMAL_min, WARNING_min)
        "basis": "Namdari S et al. (2012) — ADL 어깨 굴곡 실측",
        "grades": {
            "NORMAL":   {
                "clinical_meaning": "머리 위 선반 포함 대부분의 일상 동작 가능합니다.",
                "adl_ok":      ["식사", "머리 빗기", "세탁물 넣기", "선반 위 물건 꺼내기"],
                "adl_limited": [],
                "recommendation": "현재 어깨 가동성을 유지하세요.",
            },
            "WARNING":  {
                "clinical_meaning": "식사·세수는 가능하나 머리 위 물건 꺼내기·세탁물 넣기가 어렵습니다.",
                "adl_ok":      ["식사", "세수"],
                "adl_limited": ["머리 빗기", "세탁물 넣기", "선반 위 물건 꺼내기"],
                "recommendation": "어깨 굴곡 스트레칭 및 근력 운동을 권장합니다.",
            },
            "CRITICAL": {
                "clinical_meaning": "기본 식사 이외 대부분의 어깨 동작이 어렵습니다.",
                "adl_ok":      [],
                "adl_limited": ["식사", "머리 빗기", "세탁물 넣기", "선반 위 물건 꺼내기"],
                "recommendation": "전문 물리치료사 상담을 즉각 권고합니다.",
            },
        },
    },

    # ── 어깨 신전 (AMA 기준 40°, ADL 전용 논문 제한적) ─────────────────────
    "shoulder_extension": {
        "direction": "higher_is_better",
        "thresholds": (30.0, 15.0),
        "basis": "AMA 정상값(40°) 기반 비율 등분. ADL 전용 확정 논문 없음 — 향후 보정 권장.",
        "grades": {
            "NORMAL":   {
                "clinical_meaning": "등 뒤로 손 뻗기·재킷 입기 등 일상 동작 가능합니다.",
                "adl_ok":      ["등 뒤로 손 뻗기", "재킷 입기", "서랍 뒤로 밀기"],
                "adl_limited": [],
                "recommendation": "현재 어깨 신전 가동성을 유지하세요.",
            },
            "WARNING":  {
                "clinical_meaning": "등 뒤 동작이 제한됩니다. 옷 입기·뒤 손 뻗기가 어렵습니다.",
                "adl_ok":      ["간단한 뒤쪽 동작"],
                "adl_limited": ["등 뒤로 손 뻗기", "재킷 입기"],
                "recommendation": "어깨 신전 스트레칭을 권장합니다.",
            },
            "CRITICAL": {
                "clinical_meaning": "어깨 뒤쪽 동작이 거의 불가합니다.",
                "adl_ok":      [],
                "adl_limited": ["등 뒤 모든 동작"],
                "recommendation": "전문 물리치료사 상담을 권고합니다.",
            },
        },
    },

    # ── 어깨 외전 (Namdari 2012) ──────────────────────────────────────────
    "shoulder_abduction": {
        "direction": "higher_is_better",
        "thresholds": (120.0, 80.0),
        "basis": "Namdari S et al. (2012) — ADL 어깨 외전 실측",
        "grades": {
            "NORMAL":   {
                "clinical_meaning": "팔을 옆으로 드는 동작 대부분이 가능합니다.",
                "adl_ok":      ["식사", "머리 빗기", "세탁물 넣기", "선반 위 물건 꺼내기"],
                "adl_limited": [],
                "recommendation": "현재 어깨 외전 가동성을 유지하세요.",
            },
            "WARNING":  {
                "clinical_meaning": "식사는 가능하나 머리 빗기·세탁물 넣기가 어렵습니다.",
                "adl_ok":      ["식사"],
                "adl_limited": ["머리 빗기", "세탁물 넣기", "선반 위 물건 꺼내기"],
                "recommendation": "어깨 외전 스트레칭 및 근력 운동을 권장합니다.",
            },
            "CRITICAL": {
                "clinical_meaning": "기본 위생 동작이 어렵습니다.",
                "adl_ok":      [],
                "adl_limited": ["식사", "머리 빗기", "세탁물 넣기", "선반 위 물건 꺼내기"],
                "recommendation": "전문 물리치료사 상담을 즉각 권고합니다.",
            },
        },
    },

    # ── 팔꿈치 굴곡 (Morrey 1981) ─────────────────────────────────────────
    "elbow_flexion": {
        "direction": "higher_is_better",
        "thresholds": (130.0, 100.0),
        "basis": "Morrey BF et al. (1981) — 기능적 호 30°~130°",
        "grades": {
            "NORMAL":   {
                "clinical_meaning": "식사·세수·전화 등 기본 일상생활 동작 모두 가능합니다.",
                "adl_ok":      ["식사(포크·숟가락)", "음료 마시기", "전화 받기", "세수", "독서"],
                "adl_limited": [],
                "recommendation": "현재 팔꿈치 가동성을 유지하세요.",
            },
            "WARNING":  {
                "clinical_meaning": "전화 받기·머리 빗기가 어렵습니다. 식사는 가능합니다.",
                "adl_ok":      ["식사", "세수", "독서"],
                "adl_limited": ["전화 받기", "머리 빗기"],
                "recommendation": "팔꿈치 굴곡 스트레칭을 권장합니다.",
            },
            "CRITICAL": {
                "clinical_meaning": "식사 포함 대부분의 일상 동작이 어렵습니다.",
                "adl_ok":      [],
                "adl_limited": ["식사", "음료 마시기", "전화 받기", "세수"],
                "recommendation": "전문 물리치료사 상담을 즉각 권고합니다.",
            },
        },
    },

    # ── 팔꿈치 신전 (Morrey 1981 기능적 호 하한 30° 역산 + Orthobullets/Healio 임상 리뷰) ────────────
    "elbow_extension": {
        "direction": "lower_is_better",   # 구축각(잔류 굴곡)이 낮을수록 좋음
        "thresholds": (15.0, 30.0),       # (NORMAL_max, WARNING_max) — 구축각 기준
        # 계산: CRITICAL 경계 30° = Morrey 기능적 호 하한(30°)에 직접 적용
        #        NORMAL 경계 15° = 30° ÷ 2 (임상적 안전구역 + MediaPipe 오차 ±3~5° 마진)
        "basis": "Morrey 1981 기능적 호 하한(30°) 역산 + Orthobullets/Healio: \"30° 미만 신전 결손 대부분 환자에서 잘 견딤\"",
        "grades": {
            "NORMAL":   {
                "clinical_meaning": "팔을 거의 완전히 펼 수 있습니다. 정상 범위입니다.",
                "adl_ok":      ["물건 밀기", "팔 뻗어 잡기", "지팡이 짚기"],
                "adl_limited": [],
                "recommendation": "현재 팔꿈치 신전 가동성을 유지하세요.",
            },
            "WARNING":  {
                "clinical_meaning": "팔이 완전히 펴지지 않습니다. Morrey 기능적 호 하한에 근접합니다.",
                "adl_ok":      ["기본 굴곡 동작"],
                "adl_limited": ["완전 신전이 필요한 동작", "물건 밀기"],
                "recommendation": "팔꿈치 신전 스트레칭을 권장합니다.",
            },
            "CRITICAL": {
                "clinical_meaning": "팔꿈치 굴곡 구축(30° 이상)으로 Morrey 기능적 호 진입이 불가합니다.",
                "adl_ok":      [],
                "adl_limited": ["팔꿈치 신전 포함 대부분"],
                "recommendation": "전문 물리치료사 상담을 즉각 권고합니다.",
            },
        },
    },

    # ── 무릎 굴곡 (Rowe 2000) ────────────────────────────────────────────
    "knee_flexion": {
        "direction": "higher_is_better",
        "thresholds": (120.0, 90.0),
        "basis": "Rowe PJ et al. (2000) — ADL 무릎 굴곡 실측. 재활 목표 110°.",
        "grades": {
            "NORMAL":   {
                "clinical_meaning": "계단 오르내리기·의자 기립·보행 모두 가능합니다.",
                "adl_ok":      ["평지 보행", "계단 오르기", "의자에서 일어서기"],
                "adl_limited": [],
                "recommendation": "현재 무릎 가동성을 유지하세요.",
            },
            "WARNING":  {
                "clinical_meaning": "평지 보행은 가능하나 계단 오르내리기·의자에서 일어서기가 어렵습니다.",
                "adl_ok":      ["평지 보행"],
                "adl_limited": ["계단 오르기", "의자에서 일어서기"],
                "recommendation": "무릎 굴곡 근력 운동 및 스트레칭을 권장합니다.",
            },
            "CRITICAL": {
                "clinical_meaning": "계단·기립이 모두 불가합니다. 보행 장애 위험이 있습니다.",
                "adl_ok":      [],
                "adl_limited": ["평지 보행", "계단 오르기", "의자에서 일어서기"],
                "recommendation": "전문 물리치료사 상담을 즉각 권고합니다.",
            },
        },
    },

    # ── 무릎 신전 (Waters & Mulroy 1999 + PubMed 무릎 구축 시뮬레이션) ─────────────
    "knee_extension": {
        "direction": "lower_is_better",   # 구축각(잔류 굴곡)이 낮을수록 좋음
        "thresholds": (10.0, 15.0),       # (NORMAL_max, WARNING_max) — 구축각 기준
        # v1.1 업데이트: 기존 (10.0, 20.0) → (10.0, 15.0)
        # 계산 근거:
        #   CRITICAL 경계 15° = 건강 성인 시뮬레이션에서 15° 초과 시 VO₂ 유의 증가 (PubMed)
        #   WARNING 경계 10° = ACLR 신전 결손 5°+ 기능 저하 + MediaPipe 오차 안전 여유
        "basis": "Waters & Mulroy (1999) Gait & Posture 9(3):207-231 + PubMed: 15° 초과 시 VO₂ 유의 증가 (건강 성인 구축 시뮬레이션)",
        "grades": {
            "NORMAL":   {
                "clinical_meaning": "무릎을 거의 완전히 펼 수 있습니다. 보행 정상입니다.",
                "adl_ok":      ["평지 보행", "서 있기", "계단 내려가기"],
                "adl_limited": [],
                "recommendation": "현재 무릎 신전 가동성을 유지하세요.",
            },
            "WARNING":  {
                "clinical_meaning": "무릎이 완전히 펴지지 않습니다. 에너지 소모 증가 경계 접근. 장거리 보행 주의.",
                "adl_ok":      ["단거리 보행"],
                "adl_limited": ["장거리 보행", "계단 내려가기"],
                "recommendation": "무릎 신전 스트레칭을 권장합니다.",
            },
            "CRITICAL": {
                "clinical_meaning": "15° 이상 굴곡 구축으로 보행 에너지 소모가 유의하게 증가합니다.",
                "adl_ok":      [],
                "adl_limited": ["보행", "서 있기", "계단 내려가기"],
                "recommendation": "전문 물리치료사 상담을 즉각 권고합니다.",
            },
        },
    },
}

# 좌우 비대칭 경고 기준 (단위: 도)
ASYMMETRY_ALERT_DEG: float = 15.0


# ─────────────────────────────────────────────────────────────────────────────
# 지원 관절 목록 (외부 참조용)
# ─────────────────────────────────────────────────────────────────────────────
SUPPORTED_JOINTS: list[str] = list(_JOINTS.keys())
# ['shoulder_flexion', 'shoulder_extension', 'shoulder_abduction',
#  'elbow_flexion', 'elbow_extension', 'knee_flexion', 'knee_extension']


# ─────────────────────────────────────────────────────────────────────────────
# 핵심 함수
# ─────────────────────────────────────────────────────────────────────────────

def grade_single_rom(joint: str, rom_deg: float | None) -> Grade:
    """
    단일 ROM 수치를 NORMAL / WARNING / CRITICAL 등급으로 변환합니다.

    Args:
        joint:   관절 키 (예: "knee_flexion", "elbow_extension")
        rom_deg: 측정된 ROM 각도.
                 - 굴곡 계열(higher_is_better): 측정된 최대 굴곡각
                 - 신전 계열(lower_is_better):  잔류 굴곡각(구축각). 0°=완전신전.
                 None이면 CRITICAL 처리.

    Returns:
        "NORMAL" | "WARNING" | "CRITICAL"
    """
    if rom_deg is None or rom_deg < 0:
        return "CRITICAL"

    joint_def = _JOINTS.get(joint)
    if joint_def is None:
        print(f"[WARN] mobility_score: 미지원 관절 '{joint}'. SUPPORTED_JOINTS: {SUPPORTED_JOINTS}")
        return "NORMAL"   # 알 수 없는 관절은 판정 유보

    direction = joint_def["direction"]
    t1, t2    = joint_def["thresholds"]

    if direction == "higher_is_better":
        normal_min, warning_min = t1, t2
        if rom_deg >= normal_min:
            return "NORMAL"
        elif rom_deg >= warning_min:
            return "WARNING"
        else:
            return "CRITICAL"

    else:  # lower_is_better (신전 계열 — 구축각)
        normal_max, warning_max = t1, t2
        if rom_deg <= normal_max:
            return "NORMAL"
        elif rom_deg <= warning_max:
            return "WARNING"
        else:
            return "CRITICAL"


def _worst_grade(left: Grade, right: Grade) -> Grade:
    """좌/우 중 더 나쁜 등급을 반환합니다."""
    priority = {"CRITICAL": 0, "WARNING": 1, "NORMAL": 2}
    return left if priority[left] <= priority[right] else right


def calculate_mobility_score(rom_result: dict) -> dict:
    """
    snapshot_rom_pipeline.py의 ROM 결과를 받아 Mobility Score를 계산합니다.

    Args:
        rom_result: 최소 필요 필드:
            - "joint"     : str   (예: "shoulder_flexion")
            - "left_rom"  : float | None
            - "right_rom" : float | None

    Returns:
        Mobility Score 딕셔너리 (모듈 docstring의 [출력 스키마] 참조)
    """
    joint     = rom_result.get("joint", "unknown")
    left_rom  = _as_float(rom_result.get("left_rom"))
    right_rom = _as_float(rom_result.get("right_rom"))

    left_grade  = grade_single_rom(joint, left_rom)
    right_grade = grade_single_rom(joint, right_rom)
    overall     = _worst_grade(left_grade, right_grade)

    # 좌우 비대칭
    if left_rom is not None and right_rom is not None:
        asym_deg   = round(abs(left_rom - right_rom), 1)
        asym_alert = asym_deg >= ASYMMETRY_ALERT_DEG
    else:
        asym_deg   = None
        asym_alert = False

    # 등급별 텍스트 가져오기
    joint_def   = _JOINTS.get(joint, {})
    grade_info  = joint_def.get("grades", {}).get(overall, {})

    clinical_meaning = grade_info.get("clinical_meaning", "알 수 없는 관절입니다.")
    recommendation   = grade_info.get("recommendation", "전문가 상담을 권장합니다.")
    adl_ok           = grade_info.get("adl_ok", [])
    adl_limited      = grade_info.get("adl_limited", [])
    scoring_basis    = joint_def.get("basis", "논문 미지정")

    # 비대칭 경고 추가
    if asym_alert:
        recommendation += (
            f" 또한 좌우 차이({asym_deg}°)가 {ASYMMETRY_ALERT_DEG}° 이상으로, "
            "보행 불균형 및 부상 위험이 있습니다. 전문가 평가를 권장합니다."
        )

    return {
        "joint":            joint,
        "left_rom":         left_rom,
        "right_rom":        right_rom,
        "left_grade":       left_grade,
        "right_grade":      right_grade,
        "overall_grade":    overall,
        "asymmetry_deg":    asym_deg,
        "asymmetry_alert":  asym_alert,
        "clinical_meaning": clinical_meaning,
        "recommendation":   recommendation,
        "adl_ok":           adl_ok,
        "adl_limited":      adl_limited,
        "scoring_basis":    scoring_basis,
    }


def calculate_full_session_mobility(session_results: list[dict]) -> dict:
    """
    한 세션의 여러 관절 ROM 결과를 받아 전체 가동성 요약을 반환합니다.

    Args:
        session_results: ROM 결과 딕셔너리 리스트
            각 항목은 calculate_mobility_score()의 입력 형식과 동일

    Returns:
        세션 전체 요약 딕셔너리
    """
    joint_scores: list[dict] = [calculate_mobility_score(r) for r in session_results]

    critical = [s["joint"] for s in joint_scores if s["overall_grade"] == "CRITICAL"]
    warning  = [s["joint"] for s in joint_scores if s["overall_grade"] == "WARNING"]
    normal   = [s["joint"] for s in joint_scores if s["overall_grade"] == "NORMAL"]

    if critical:
        session_overall = "CRITICAL"
        summary = (
            f"위험 관절({', '.join(critical)})이 발견되었습니다. "
            "전문 물리치료사 상담을 즉각 권고합니다."
        )
    elif warning:
        session_overall = "WARNING"
        summary = (
            f"주의 관절({', '.join(warning)})이 있습니다. "
            "해당 관절 스트레칭 및 근력 운동을 권장합니다."
        )
    else:
        session_overall = "NORMAL"
        summary = "모든 측정 관절의 가동성이 일상생활 기능 범위 내에 있습니다."

    return {
        "session_overall_grade": session_overall,
        "critical_joints":       critical,
        "warning_joints":        warning,
        "normal_joints":         normal,
        "joint_scores":          joint_scores,
        "summary_text":          summary,
    }


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼
# ─────────────────────────────────────────────────────────────────────────────

def _as_float(v) -> float | None:
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ─────────────────────────────────────────────────────────────────────────────
# 독립 실행 테스트
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    print("\n[Tro-Fit] Mobility Score Engine v1.1 — Self Test")
    print(f"  지원 관절: {SUPPORTED_JOINTS}")
    print("=" * 65)

    # ── 테스트 케이스 ─────────────────────────────────────────────────────
    test_cases = [
        # (설명,          joint,                left_rom, right_rom)
        ("어깨굴곡 NORMAL",  "shoulder_flexion",   135.0,    128.0),
        ("어깨외전 WARNING", "shoulder_abduction", 95.0,     110.0),
        ("팔꿈치굴곡 WARN",  "elbow_flexion",      105.0,    88.0),
        ("팔꿈치신전 NORM",  "elbow_extension",    8.0,      12.0),   # 구축각 8°/12°
        ("팔꿈치신전 CRIT",  "elbow_extension",    35.0,     28.0),   # 구축각 35°/28°
        ("무릎굴곡 WARNING", "knee_flexion",       115.0,    98.0),
        ("무릎신전 NORMAL",  "knee_extension",     5.0,      8.0),    # 구축각 5°/8°
        ("비대칭 경고",      "shoulder_flexion",   135.0,    95.0),   # 40° 차이
    ]

    for desc, joint, left, right in test_cases:
        r = calculate_mobility_score({"joint": joint, "left_rom": left, "right_rom": right})
        print(f"\n  [{desc}]")
        print(f"    관절     : {r['joint']}")
        print(f"    좌/우    : {r['left_rom']}° ({r['left_grade']}) / {r['right_rom']}° ({r['right_grade']})")
        print(f"    전체등급  : {r['overall_grade']}")
        if r['asymmetry_alert']:
            print(f"    ⚠ 비대칭  : {r['asymmetry_deg']}°")
        print(f"    임상의미  : {r['clinical_meaning']}")
        print(f"    가능ADL  : {r['adl_ok'] or '없음'}")
        print(f"    제한ADL  : {r['adl_limited'] or '없음'}")

    # ── 세션 전체 테스트 ──────────────────────────────────────────────────
    print("\n\n  [전체 세션 테스트]")
    print("-" * 65)
    session = [
        {"joint": "shoulder_flexion",   "left_rom": 135.0, "right_rom": 128.0},
        {"joint": "shoulder_abduction", "left_rom": 95.0,  "right_rom": 110.0},
        {"joint": "elbow_flexion",      "left_rom": 105.0, "right_rom": 88.0},
        {"joint": "knee_flexion",       "left_rom": 115.0, "right_rom": 98.0},
    ]
    sr = calculate_full_session_mobility(session)
    print(f"\n  세션 전체 등급: {sr['session_overall_grade']}")
    print(f"  위험 관절    : {sr['critical_joints'] or '없음'}")
    print(f"  주의 관절    : {sr['warning_joints'] or '없음'}")
    print(f"  정상 관절    : {sr['normal_joints'] or '없음'}")
    print(f"\n  요약: {sr['summary_text']}")
    print()
