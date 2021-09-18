
class ParseDate:
    def parse_date(self, str):
        substr = str[0:10]
        dt_list = substr.split('-')
        date = f"{dt_list[2]}.{dt_list[1]}.{dt_list[0]}"
        return date

# pd = ParseDate()
# print(pd.parse_date("2021-09-17 12:47:32+00:00"))
