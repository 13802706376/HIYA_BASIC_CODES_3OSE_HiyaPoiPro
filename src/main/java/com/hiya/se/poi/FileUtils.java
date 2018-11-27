package com.hiya.se.poi;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class FileUtils<T>
{
    public static void close(Object o,String methodName)
    {
        if (null != o)
        {
            Method closeMethod = null;
            try
            {
                 closeMethod = o.getClass().getDeclaredMethod(methodName);
                 closeMethod.invoke(o, methodName);
            } catch (NoSuchMethodException | SecurityException | IllegalAccessException | InvocationTargetException e)
            {
                e.printStackTrace();
            }
        }
    }
}
