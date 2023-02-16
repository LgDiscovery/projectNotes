package com.lg.project.notes.common.dataStructure;

import java.util.*;

/**
 * 求中值
 * @param <E>
 */
public class Median<E> {
    private PriorityQueue<E> minP; //小堆
    private PriorityQueue<E> maxP; //大堆
    private E m; // 中位数

    public Median() {
        this.minP = new PriorityQueue<>();
        this.maxP = new PriorityQueue<>(11, Collections.reverseOrder());
    }

    private int compare(E e, E m) {
        Comparable<? super E> cmpr = (Comparable<? super E>) e;
        return cmpr.compareTo(m);
    }

    public void add(E e) {
        if (m == null) { //第一个元素
            m = e;
            return;
        }
        if (compare(e, m) <= 0) {
            //小于中值，加入最大堆中
            maxP.add(e);
        } else {
            minP.add(e);
        }
        if (minP.size() - maxP.size() >= 2) {
            maxP.add(this.m);
            this.m = minP.poll();
        } else if (maxP.size() - minP.size() >= 2) {
            minP.add(this.m);
            this.m = maxP.poll();
        }
    }

    public void addAll(Collection<? extends E> c) {
        for (E e : c) {
            add(e);
        }
    }

    public E getM() {
        return m;
    }

    public static void main(String[] args) {
        Median<Integer> median = new Median<>();
        List<Integer> list = Arrays.asList(new Integer[]{
                34, 90, 67, 45, 1, 4, 5, 6, 7, 9, 10
        });
        median.addAll(list);
        System.out.println(median.getM());
    }
}