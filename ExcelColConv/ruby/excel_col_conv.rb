class ExcelColConv
    AZ_LENGTH = 'Z'.ord - 'A'.ord + 1 #26
    OFFSET_NUM = 'A'.ord - 1 #64

    def to_col_number(col_str)
       convert_str 0, col_str.split('') 
    end

    def convert_str(ret, str_array)
        if str_array.empty?
            ret
        else
            c = str_array.shift
            convert_str calc_decimal(c, str_array.length) + ret, str_array
        end
    end
    
    def calc_decimal(c, times)
        AZ_LENGTH ** times * (c.ord - OFFSET_NUM)
    end

    def to_col_string(col_num)
        convert_num '', col_num 
    end

    def convert_num(ret, n)
        if n == 0
            ret
        else
            quo, rem = excel_divmod n
            convert_num (rem + OFFSET_NUM).chr + ret, quo            
        end
    end    

    def excel_divmod(n)
        quo, rem = n.divmod AZ_LENGTH
        rem == 0 ? [quo - 1, AZ_LENGTH] : [quo, rem]
    end
end
