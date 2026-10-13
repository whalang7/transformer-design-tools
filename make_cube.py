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
from win32com.client import gencache

# ---------------------------------------------------------------
# 치수 설정 (단위: 미터)  ※ 100 mm = 0.1 m
# ---------------------------------------------------------------
SIDE = 0.1    # 정육면체 한 변의 길이 (100 mm)
HALF = SIDE / 2  # 원점 기준으로 사각형을 그리기 위한 절반 길이 (50 mm)

# Solid Edge 상수 타입 라이브러리(SolidEdgeConstants)의 GUID와 버전.
# [확인 필요] 설치된 Solid Edge 버전에 따라 GUID/버전이 다를 수 있다.
#   확인 방법: python -m win32com.client.makepy 실행 후 목록에서
#   "Solid Edge Constants Type Library" 항목을 찾는다.
SE_CONSTANTS_TLB = ("{C467A6F5-27ED-11D2-BE30-080036B4D502}", 0, 1, 0)


def fail(step, err):
    """실패한 단계 이름과 오류 내용을 한국어로 출력하고 프로그램을 끝낸다."""
    print(f"❌ [{step}] 단계에서 실패했습니다: {err}")
    sys.exit(1)


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
        # 이 작업이 끝나야 win32com.client.constants 에서 이름으로 꺼낼 수 있다.
        gencache.EnsureModule(*SE_CONSTANTS_TLB)
        c = win32com.client.constants
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
        # [확인 필요] igLineStart / igLineEnd 상수명 (KeypointIndexConstants)
        rel = profile.Relations2d
        rel.AddKeypoint(l1, c.igLineEnd, l2, c.igLineStart)
        rel.AddKeypoint(l2, c.igLineEnd, l3, c.igLineStart)
        rel.AddKeypoint(l3, c.igLineEnd, l4, c.igLineStart)
        rel.AddKeypoint(l4, c.igLineEnd, l1, c.igLineStart)

        # 프로파일을 닫고 검사한다. 0이 아니면 닫힌 도형이 아니라는 뜻.
        # [확인 필요] igProfileClosed 상수명 (ProfileValidationType)
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
        # [확인 필요] igRight 상수명 (FeaturePropertyConstants). 방향이 반대로 나오면 igLeft 사용
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
    main()
