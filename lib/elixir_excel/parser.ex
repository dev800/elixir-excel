defmodule ElixirExcel.Parser do

  @start_time ~N[1899-12-30 00:00:00.000000]
  # 1904年系统的excel的起始日期
  @start_time1904 ~N[1904-01-01 00:00:00.000000]

  def parse_time(value, is_1904 \\ false)

  def parse_time(value, true) when is_float(value) do
    Timex.shift(@start_time1904, microseconds: TinyUtil.Integer.parse(value * 86_400_000_000))
  end

  def parse_time(value, false) when is_float(value) do
    Timex.shift(@start_time, microseconds: TinyUtil.Integer.parse(value * 86_400_000_000))
  end

  def parse_time(value, true) when is_integer(value) do
    Timex.shift(@start_time1904, days: value)
  end

  def parse_time(value, false) when is_integer(value) do
    Timex.shift(@start_time, days: value)
  end
end
