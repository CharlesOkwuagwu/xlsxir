defmodule XlsxirTest do
  use ExUnit.Case

  import Xlsxir

  def path(), do: "./test/test_data/test.xlsx"

  test "second worksheet is parsed with index argument of 1" do
    {:ok, {pid, table_id}} = extract(path(), 1)
    assert get_list(table_id) == [[1, 2], [3, 4]]
    close(pid)
  end

  test "able to parse maximum number of columns" do
    {:ok, {pid, table_id}} = extract(path(), 2)
    assert get_cell(table_id, "XFD1") == 16384
    close(pid)
  end

  test "able to parse maximum number of rows" do
    {:ok, {pid, table_id}} = extract(path(), 3)
    assert get_cell(table_id, "A1048576") == 1048576
    close(pid)
  end

  test "able to parse cells with errors" do
    {:ok, {pid, table_id}} = extract(path(), 4)
    assert get_list(table_id) == [["#DIV/0!", "#REF!", "#NUM!", "#VALUE!"]]
    close(pid)
  end

  test "able to parse custom formats" do
    {:ok, {pid, table_id}} = extract(path(), 5)
    assert get_list(table_id) == [[-123.45, 67.89, {2015, 1, 1}, {2016, 12, 31}, {15, 12, 45}, ~N[2012-12-18 14:26:00]]]
    close(pid)
  end

  test "able to parse with conditional formatting" do
    {:ok, {pid, table_id}} = extract(path(), 6)
    assert get_list(table_id) == [["Conditional"]]
    close(pid)
  end

  test "able to parse with boolean values" do
    {:ok, {pid, table_id}} = extract(path(), 7)
    assert get_list(table_id) == [[true, false]]
    close(pid)
  end

  test "peek file contents" do
    {:ok, {pid, table_id}} = peek(path(), 8, 10)
    assert get_cell(table_id, "G10") == 8437
    assert get_info(table_id, :rows) == 10
    close(pid)
  end
end
