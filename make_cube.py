# -*- coding: utf-8 -*-
"""
Solid Edge를 띄우고 새 파트 문서에 100 x 100 x 100 mm 정육면체를 만드는 스크립트.

실행: python make_cube.py
필요: Windows, Solid Edge 2024 이상, Python 3.x, pywin32
주의: Solid Edge API의 길이 단위는 '미터'이므로 100 mm = 0.1 로 입력한다.
"""

import sys

import pythoncom
import win32com.client
from win32com.client import gencache, selecttlb

# ---------------------------------------------------------------
# 치수 설정 (단위: 미터)  ※ 100 mm = 0.1 m
# ---------------------------------------------------------------
SIDE = 0.1    # 정육면체 한 변의 길이 (100 mm)
HALF = SIDE / 2  # 원점 기준으로 사각형을 그리기 위한 절반 길이 (50 mm)

# Solid Edge 상수 타입 라이브러리의 이름 (Siemens API 문서에 표기된 이름).
# GUID를 직접 적지 않고, 이 이름으로 Windows에 등록된 라이브러리를 찾는다.
SE_CONSTANTS_NAME = "Solid Edge Constants Type Library"


def fail(step, err):
    """실패한 단계 이름과 오류 내용을 한국어로 출력하고 프로그램을 끝낸다."""
    print(f"❌ [{step}] 단계에서 실패했습니다: {err}")
    sys.exit(1)


def load_constants():
    """상수 타입 라이브러리를 찾아 불러오고, 상수 묶음(constants)을 돌려준다."""
    # Windows에 등록된 타입 라이브러리 중 이름이 일치하는 것을 찾는다
    specs = selecttlb.FindTlbsWithDescription(SE_CONSTANTS_NAME)
    if not specs:
        raise RuntimeError(f"'{SE_CONSTANTS_NAME}'를 찾을 수 없습니다")
    # 찾은 파일에서 GUID와 버전을 읽어 파이썬용 상수 모듈을 만든다
    tlb = pythoncom.LoadTypeLib(specs[-1].dll)
    guid, lcid, _, major, minor, _ = tlb.GetLibAttr()
    gencache.EnsureModule(str(guid), lcid, major, minor)
    return win32com.client.constants


def main():
    # -----------------------------------------------------------
    # 1단계: Solid Edge 연결
    #   이미 켜져 있으면 그 프로그램에 연결하고, 없으면 새로 실행한다.
    # -----------------------------------------------------------
    try:
        try:
            # 실행 중인 Solid Edge 찾기
            running = win32com.client.GetActiveObject("SolidEdge.Application")
            # 타입 정보(메서드 목록 등)를 불러와 쓰기 편한 객체로 감싼다
            app = gencache.EnsureDispatch(running)
            print("실행 중인 Solid Edge에 연결했습니다.")
        except pythoncom.com_error:
            # 실행 중인 것이 없으면 새로 실행
            app = gencache.EnsureDispatch("SolidEdge.Application")
            print("Solid Edge를 새로 실행했습니다.")
        app.Visible = True  # 화면에 보이게 한다

        # 상수(igProfileClosed 등)를 쓰기 위해 상수 타입 라이브러리를 불러온다.
        # 이 작업이 끝나야 c.igProfileClosed 처럼 이름으로 꺼내 쓸 수 있다.
        c = load_constants()
    except Exception as err:
        fail("Solid Edge 연결", err)
    print("✅ Solid Edge 연결 완료")

    # -----------------------------------------------------------
    # 2단계: 새 파트(.par) 문서 만들기 (기본 ISO 템플릿 사용)
    # -----------------------------------------------------------
    try:
        doc = app.Documents.Add("SolidEdge.PartDocument")
        # 파트 문서 전용 메서드(RefPlanes, Models 등)를 쓰기 위해 타입 정보를 불러온다
        part = gencache.EnsureDispatch(doc)
    except Exception as err:
        fail("새 파트 문서 생성", err)
    print("✅ 새 파트 문서 생성 완료")

    # -----------------------------------------------------------
    # 3단계: 스케치 - 윗면(XY) 평면에 원점 중심 100 x 100 mm 사각형 그리기
    # -----------------------------------------------------------
    try:
        # 기준평면 1번 = Top(XY) 평면
        # [확인 필요] 템플릿에 따라 1번이 Top 평면이 아닐 수 있음
        ref_plane = part.RefPlanes.Item(1)
        print(f"사용한 기준평면: {ref_plane.Name}")  # Top(윗면)인지 눈으로 확인용

        # 프로파일(스케치) 묶음을 만들고, 그 안에 기준평면 위의 프로파일을 추가
        profile_set = part.ProfileSets.Add()
        profile = profile_set.Profiles.Add(ref_plane)

        # 사각형의 네 꼭짓점 (x, y) - 단위는 미터
        p1 = (-HALF, -HALF)  # 왼쪽 아래
        p2 = (HALF, -HALF)   # 오른쪽 아래
        p3 = (HALF, HALF)    # 오른쪽 위
        p4 = (-HALF, HALF)   # 왼쪽 위

        # 직선 4개로 사각형 그리기 (시작점 x, y, 끝점 x, y)
        lines = profile.Lines2d
        l1 = lines.AddBy2Points(p1[0], p1[1], p2[0], p2[1])
        l2 = lines.AddBy2Points(p2[0], p2[1], p3[0], p3[1])
        l3 = lines.AddBy2Points(p3[0], p3[1], p4[0], p4[1])
        l4 = lines.AddBy2Points(p4[0], p4[1], p1[0], p1[1])

        # 각 선의 끝점과 다음 선의 시작점을 붙여서(키포인트 구속) 닫힌 도형으로 만든다
        # [확인 필요] igLineStart / igLineEnd 상수명 (공개 예제에서 사용 확인, 공식 문서 미확인)
        rel = profile.Relations2d
        rel.AddKeypoint(l1, c.igLineEnd, l2, c.igLineStart)
        rel.AddKeypoint(l2, c.igLineEnd, l3, c.igLineStart)
        rel.AddKeypoint(l3, c.igLineEnd, l4, c.igLineStart)
        rel.AddKeypoint(l4, c.igLineEnd, l1, c.igLineStart)

        # 프로파일을 닫고 검사한다. 0이 아니면 닫힌 도형이 아니라는 뜻.
        status = profile.End(c.igProfileClosed)
        if status != 0:
            raise RuntimeError(f"프로파일이 닫히지 않았습니다 (반환값 {status})")
    except Exception as err:
        fail("스케치(사각형 프로파일)", err)
    print("✅ 100 x 100 mm 사각형 스케치 완료")

    # -----------------------------------------------------------
    # 4단계: 돌출 - 사각형을 100 mm(0.1 m) 밀어 올려 정육면체 만들기
    # -----------------------------------------------------------
    try:
        # 인자: 프로파일 개수, 프로파일 목록, 돌출 방향, 돌출 거리(미터)
        # 방향이 반대로 나오면 igRight 대신 igLeft 사용
        part.Models.AddFiniteExtrudedProtrusion(1, [profile], c.igRight, SIDE)
        profile.Visible = False  # 다 쓴 스케치 선은 숨긴다
    except Exception as err:
        fail("돌출", err)
    print("✅ 100 mm 돌출(정육면체 생성) 완료")

    # -----------------------------------------------------------
    # 5단계: 화면 맞춤(Fit). 파일은 저장하지 않는다.
    # -----------------------------------------------------------
    try:
        app.ActiveWindow.View.Fit()
    except Exception as err:
        fail("화면 맞춤", err)
    print("✅ 화면 맞춤 완료 (파일은 저장하지 않았습니다)")


if __name__ == "__main__":
    # 출력이 파일/파이프로 갈 때 ✅ 같은 문자 때문에 오류가 나지 않도록 UTF-8로 맞춘다
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    try:
        main()
    finally:
        # 실행파일(.exe)을 더블클릭하면 창이 바로 닫히므로, 결과를 읽을 시간을 준다
        if getattr(sys, "frozen", False):
            try:
                input("엔터를 누르면 창이 닫힙니다...")
            except EOFError:
                pass
