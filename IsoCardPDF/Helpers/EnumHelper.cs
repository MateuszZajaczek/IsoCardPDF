using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

public static class EnumHelper
{
    public static bool TryParseDescription<TEnum>(string description, out TEnum result) where TEnum : struct, IConvertible
    {
        if (!typeof(TEnum).IsEnum)
        {
            throw new ArgumentException("TEnum must be an enumerated type");
        }

        foreach (var field in typeof(TEnum).GetFields())
        {
            var attribute = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute;

            if (attribute != null)
            {
                if (attribute.Description == description)
                {
                    result = (TEnum)field.GetValue(null);
                    return true;
                }
            }
            else
            {
                if (field.Name == description)
                {
                    result = (TEnum)field.GetValue(null);
                    return true;
                }
            }
        }

        result = default(TEnum);
        return false;
    }
}
