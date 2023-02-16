package com.lg.project.notes.common.dataStructure;

import java.util.Arrays;
import java.util.Collection;
import java.util.PriorityQueue;

/**
 * 求K个最大的元素
 * @param <E>
 */
public class TopK<E> {
    private PriorityQueue<E> p;
    private int k;
    private Comparable<? super E> head;

    public TopK(int k) {
        this.k = k;
        this.p = new PriorityQueue<>(k);
    }

    public void addAll(Collection<? extends E> c) {
        for (E e : c) {
            add(e);
        }
    }

    public void add(E e) {
        if (p.size() < k) {
            p.add(e);
            return;
        }
        Comparable<? super E> head = (Comparable<? super E>) p.peek();
        if (head.compareTo(e) > 0) {
            return;
        }

        p.poll();
        p.add(e);
    }

    public <T> T[] toArray(T[] a) {
        return p.toArray(a);
    }

    public E getKth() {
        return p.peek();
    }

    public static void main(String[] args) {
        TopK<Integer> top5 = new TopK<>(5);
        top5.addAll(Arrays.asList(new Integer[]{
                100, 1, 2, 5, 6, 7, 34, 9, 3, 4, 5, 8, 23, 21, 90, 1, 0
        }));
        System.out.println(Arrays.toString(top5.toArray(new Integer[0])));
        System.out.println(top5.getKth());
    }
}