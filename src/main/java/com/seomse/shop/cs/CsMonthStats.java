package com.seomse.shop.cs;

import java.util.HashMap;
import java.util.Map;

/**
 * <pre>
 *  파 일 명 : CsMonthStats.java
 *  설    명 : Cs 월별 통계 데이터 유형
 *
 *  작 성 자 : macle
 *  작 성 일 : 2019.08
 *  버    전 : 1.0
 *  수정이력 :
 *  기타사항 :
 * </pre>
 * @author Copyrights 2019 by ㈜섬세한사람들. All right reserved.
 */
public class CsMonthStats {
    String ym;
    int count = 0;

    Map<String, NameCount> shopCountMap = new HashMap<>();

    //1레벨
    Map<String, CsCount> csCountMap= new HashMap<>();


}
