package com.nouseen.Service;

import com.nouseen.bean.Ant;

import java.util.*;

/**
 * Created by nouseen on 2017/9/20.
 */
public class AntService {

    // 单例
    private static AntService me = new AntService();

    public static AntService me() {
        return me;
    }

    private AntService(){}

    /**
     * 爬行
     * @param ant
     */
    public void run(Ant ant) {
        double position = ant.getPosition();
        if (ant.getHeadTo() == HeadTo.forwordValue) {
            ant.setPosition(position + ant.getSpeed());
        } else {
            ant.setPosition(position - ant.getSpeed());

        }



    }

    /**
     * 位置是否已出区间
     * @param ant
     * @param start
     * @param end
     * @return
     */
    public boolean isAntOut(Ant ant, double start, double end) {
        if (ant.getPosition() <= start || ant.getPosition() >= end) {
            return true;
        }

        return false;
    }

    public static void main(String[] args) {

        // 服务类
        AntService antService = AntService.me();

        // 位置区间
        double start = 0;
        double end = 27;

        // 蚂蚁列表
        List<Ant> antList = new ArrayList<Ant>();

        // 定义5个蚂蚁
        Ant ant1 = new Ant(1, 1, 3, HeadTo.forwordValue );
        Ant ant2 = new Ant(1, 1, 7, HeadTo.backValue);
        Ant ant3 = new Ant(1, 1, 11, HeadTo.backValue);
        Ant ant4 = new Ant(1, 1, 17, HeadTo.forwordValue);
        Ant ant5 = new Ant(1, 1, 23, HeadTo.backValue);

        antList.add(ant1);
        antList.add(ant2);
        antList.add(ant3);
        antList.add(ant4);
        antList.add(ant5);

        int time = 0;
        // 如果蚂蚁没都走出范围，则继续走下一步（不考虑两只蚂蚁在0.5时相遇）
        while (! antList.isEmpty()) {
            // 时间加1
            ++time;
            Iterator<Ant> antIterator = antList.iterator();
            while (antIterator.hasNext()) {
                // 拿到当前蚂蚁
                Ant ant = antIterator.next();
                // 走一步
                antService.run(ant);
                // 是否走出范围
                boolean isAntOut = antService.isAntOut(ant, start, end);
                // 如果走出范围，则移除
                if (isAntOut) {
                    // 标记离场时间
                    ant.setOutTime(time);
                    antIterator.remove();
                }
            }

            // 给蚂蚁调头
            antService.turnHeadAntAtSamePosition(antList);

        }

        System.out.println(String.format("共花了%s秒",time));
    }

    /**
     * 碰头
     * @param antList
     * @return
     */
    public void turnHeadAntAtSamePosition(List<Ant> antList) {

        AntService antService = me();
        Map<Double,Ant> antPositionMap = new HashMap<Double,Ant>();

        // 遍历所有的蚂蚁，处理其碰头业务
        for (Ant ant : antList) {

            // 拿到这里的另一只蚂蚁
            Ant otherAnt = antPositionMap.get(ant.getPosition());

            // 如果这里有一只，则把两只都调头
            if (otherAnt != null) {
                antService.turnHead(otherAnt);
                antService.turnHead(ant);
                continue;
            }

            // 如果这里没有，则在这里放一只
            antPositionMap.put(ant.getPosition(), ant);

        }

        return;
    }

    /**
     * 调头
     * @param ant
     */
    public void turnHead(Ant ant) {
        // 如果头朝前则改为朝后
        if (ant.getHeadTo() == HeadTo.forwordValue) {
            ant.setHeadTo(HeadTo.backValue);
            return;
        }

        // 改为朝前
        ant.setHeadTo(HeadTo.forwordValue);
    }

    /**
     * 方向内部类
     */
    public static class HeadTo {

        public static int forwordValue = 1;
        public static String forwordLable = "向前";

        public static int backValue = 2;
        public static String backLable = "向后";

        public static Map<Integer, String> dataMap;

        static{
            dataMap = new HashMap<Integer, String>();
            dataMap.put(backValue, backLable);
            dataMap.put(forwordValue, forwordLable);
        }

    }
}

