package com.hiya.se.poi;

import java.util.function.Supplier;
public interface IPio
{
    void doCreate(String path);
    void doParse(String path);
    public static IPio create (Supplier<IPio>  supplier)
    {
        return supplier.get();
    }
}
