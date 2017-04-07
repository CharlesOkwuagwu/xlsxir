defmodule XlsxirTest do
  use ExUnit.Case

  import Xlsxir

  def path(), do: "./test/test_data/test.xlsx"

  test "second worksheet is parsed with index argument of 1" do
    {:ok, pid} = extract(path(), 1)
    assert get_list(pid) == [[1, 2], [3, 4]]
    close(pid)
  end

  test "able to parse maximum number of columns" do
    {:ok, pid} = extract(path(), 2)
    assert get_cell(pid, "XFD1") == 16384
    close(pid)
  end

  test "able to parse maximum number of rows" do
    {:ok, pid} = extract(path(), 3)
    assert get_cell(pid, "A1048576") == 1048576
    close(pid)
  end

  test "able to parse cells with errors" do
    {:ok, pid} = extract(path(), 4)
    assert get_list(pid) == [["#DIV/0!", "#REF!", "#NUM!", "#VALUE!"]]
    close(pid)
  end

  test "able to parse custom formats" do
    {:ok, pid} = extract(path(), 5)
    assert get_list(pid) == [[-123.45, 67.89, {2015, 1, 1}, {2016, 12, 31}, {15, 12, 45}, ~N[2012-12-18 14:26:00]]]
    close(pid)
  end

  test "able to parse with conditional formatting" do
    {:ok, pid} = extract(path(), 6)
    assert get_list(pid) == [["Conditional"]]
    close(pid)
  end

  test "able to parse with boolean values" do
    {:ok, pid} = extract(path(), 7)
    assert get_list(pid) == [[true, false]]
    close(pid)
  end

  test "peek file contents" do
    {:ok, pid} = peek(path(), 8, 10)
    assert get_cell(pid, "G10") == 8437
    assert get_info(pid, :rows) == 10
    close(pid)
  end
end
