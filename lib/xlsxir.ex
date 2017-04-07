defmodule Xlsxir do
  alias Xlsxir.{SaxParser, Unzip}

  defstruct [styles: nil, shared_strings: nil, worksheets: [], max_rows: nil, timer: nil]

  use Application

  def start(_type, _args) do
    import Supervisor.Spec

    children = [
      worker(Xlsxir.StateManager, []),
    ]

    opts = [strategy: :one_for_one, name: __MODULE__]
    Supervisor.start_link(children, opts)
  end

  @moduledoc """
  Extracts and parses data from a `.xlsx` file to an Erlang Term Storage (ETS) process and provides various functions for accessing the data.
  """

  @doc """
  Extracts worksheet data contained in the specified `.xlsx` file to an ETS process. Successful extraction
  returns `{:ok, tid}` with the timer argument set to false and returns a tuple of `{:ok, tid, time}` where time is a list containing time elapsed during the extraction process
  (i.e. `[hour, minute, second, microsecond]`) when the timer argument is set to true and tid - is the ETS table id

  Cells containing formulas in the worksheet are extracted as either a `string`, `integer` or `float` depending on the resulting value of the cell.
  Cells containing an ISO 8601 date format are extracted and converted to Erlang `:calendar.date()` format (i.e. `{year, month, day}`).

  ## Parameters
  - `path` - file path of a `.xlsx` file type in `string` format
  - `index` - index of worksheet from within the Excel workbook to be parsed (zero-based index)
  - `timer` - boolean flag that tracks extraction process time and returns it when set to `true`. Default value is `false`.

  ## Example
  Extract first worksheet in an example file named `test.xlsx` located in `./test/test_data`:

        iex> {:ok, tid} = Xlsxir.extract("./test/test_data/test.xlsx", 0)
        iex> Enum.member?(:ets.all, tid)
        true
        iex> Xlsxir.close(tid)
        :ok

  ## Test parallel parsing
        iex> task1 = Task.async(fn -> Xlsxir.extract("./test/test_data/test.xlsx", 0) end)
        iex> task2 = Task.async(fn -> Xlsxir.extract("./test/test_data/test.xlsx", 0) end)
        iex> {:ok, tid1} = Task.await(task1)
        iex> {:ok, tid2} = Task.await(task2)
        iex> Xlsxir.get_list(tid1)
        [["string one", "string two", 10, 20, {2016, 1, 1}]]
        iex> Xlsxir.get_list(tid2)
        [["string one", "string two", 10, 20, {2016, 1, 1}]]
        iex> Xlsxir.close(tid1)
        :ok
        iex> Xlsxir.close(tid2)
        :ok

  """
  def extract(path, index, timer \\ false) do
    excel = if timer do
      {_, s, ms} = :erlang.timestamp
      %__MODULE__{timer: [s, ms]}
    else
      %__MODULE__{}
    end

    case Unzip.validate_path_and_index(path, index) do
      {:ok, file}      ->
        case extract_xml(file, index) do
          {:ok, file_paths} ->
            {excel, result} = do_extract(excel, file_paths, index)
            :ets.delete(excel.styles)
            :ets.delete(excel.shared_strings)
            result
          {:error, reason}  -> {:error, reason}
        end
      {:error, reason} -> {:error, reason}
    end
  end

  defp do_extract(%__MODULE__{timer: timer} = excel, file_paths, index) do
    excel = Enum.reduce(file_paths, excel, fn {file, content}, acc ->
      cond do
        file == 'xl/sharedStrings.xml' && is_nil(excel.shared_strings) ->
          {:ok, %Xlsxir.ParseString{tid: tid}, _} = SaxParser.parse(content, :string)
          %{acc | shared_strings: tid}
        file == 'xl/styles.xml' && is_nil(excel.styles) ->
          {:ok, %Xlsxir.ParseStyle{tid: tid}, _} = SaxParser.parse(content, :style)
          %{acc | styles: tid}
        true -> acc
      end
    end)

    {_, content} = Enum.find(file_paths, fn {path, _} ->
      path == 'xl/worksheets/sheet#{index + 1}.xml'
    end)

    {:ok, %Xlsxir.ParseWorksheet{tid: tid}, _} = SaxParser.parse(content, :worksheet, excel)

    if !is_nil(timer) do
      {excel, {:ok, tid, stop_timer(excel.timer)}}
    else
      {excel, {:ok, tid}}
    end
  end

  @doc """
  Extracts the first n number of rows from the specified worksheet contained in the specified `.xlsx` file to an ETS process.
  Successful extraction returns `{:ok, tid}` where tid - is ETS table id.

  ## Parameters
  - `path` - file path of a `.xlsx` file type in `string` format
  - `index` - index of worksheet from within the Excel workbook to be parsed (zero-based index)
  - `rows` - the number of rows to fetch from within the specified worksheet

  ## Example
  Peek at the first 10 rows of the 9th worksheet in an example file named `test.xlsx` located in `./test/test_data`:

        iex> {:ok, tid} = Xlsxir.peek("./test/test_data/test.xlsx", 8, 10)
        iex> Enum.member?(:ets.all, tid)
        true
        iex> Xlsxir.close(tid)
        :ok
  """
  def peek(path, index, rows) do
    excel = %__MODULE__{max_rows: rows}
    case Unzip.validate_path_and_index(path, index) do
      {:ok, file}      ->
        case extract_xml(file, index) do
          {:ok, file_paths} ->
            {excel, result} = do_extract(excel, file_paths, index)
            :ets.delete(excel.styles)
            :ets.delete(excel.shared_strings)
            result
          {:error, reason}  -> {:error, reason}
        end
      {:error, reason} -> {:error, reason}
    end
  end

  defp extract_xml(file, index) do
    Unzip.xml_file_list(index)
    |> Unzip.extract_xml_to_memory(file)
  end

  @doc """
  Extracts worksheet data contained in the specified `.xlsx` file to an ETS process. Successful extraction
  returns `{:ok, table_id}` with the timer argument set to false and returns a tuple of `{:ok, table_id, time}` where `time` is a list containing time elapsed during the extraction process
  (i.e. `[hour, minute, second, microsecond]`) when the timer argument is set to true. The `table_id` is used to access data for that particular ETS process with the various access functions of the
  `Xlsxir` module.

  Cells containing formulas in the worksheet are extracted as either a `string`, `integer` or `float` depending on the resulting value of the cell.
  Cells containing an ISO 8601 date format are extracted and converted to Erlang `:calendar.date()` format (i.e. `{year, month, day}`).

  ## Parameters
  - `path` - file path of a `.xlsx` file type in `string` format
  - `index` - index of worksheet from within the Excel workbook to be parsed (zero-based index)
  - `timer` - boolean flag that tracts extraction process time and returns it when set to `true`. Defalut value is `false`.

  ## Example
  Extract first worksheet in an example file named `test.xlsx` located in `./test/test_data`:

        iex> {:ok, tid} = Xlsxir.multi_extract("./test/test_data/test.xlsx", 0)
        iex> Enum.member?(:ets.all, tid)
        true
        iex> Xlsxir.close(tid)
        :ok

  ## Example
  Extract all worksheets in an example file named `test.xlsx` located in `./test/test_data`:

        iex> results = Xlsxir.multi_extract("./test/test_data/test.xlsx")
        iex> alive_ids = Enum.map(results, fn {:ok, tid} -> Enum.member?(:ets.all, tid) end)
        iex> Enum.all?(alive_ids)
        true
        iex> Enum.map(results, fn {:ok, tid} -> Xlsxir.close(tid) end) |> Enum.all?(fn result -> result == :ok end)
        true

  ## Example
  Extract all worksheets in an example file named `test.xlsx` located in `./test/test_data` with timer:

        iex> results = Xlsxir.multi_extract("./test/test_data/test.xlsx", nil, true)
        iex> alive_ids = Enum.map(results, fn {:ok, tid, _timer} -> Enum.member?(:ets.all, tid) end)
        iex> Enum.all?(alive_ids)
        true
        iex> Enum.map(results, fn {:ok, tid, _timer} -> Xlsxir.close(tid) end) |> Enum.all?(fn result -> result == :ok end)
        true
  """
  def multi_extract(path, index \\ nil, timer \\ nil, excel \\ nil) do
    {excel, initial_parse} = if is_nil(excel), do: {%__MODULE__{}, true}, else: {excel, false}
    case is_nil(index) do
      true ->
        case Unzip.validate_path_all_indexes(path) do
          {:ok, indexes} ->
            {excel, result} = Enum.reduce(indexes, {excel, []}, fn i, {acc_excel, acc} ->
              {new_excel, result} = multi_extract(path, i, timer, acc_excel)
              {new_excel, acc ++ [result]}
            end)
            :ets.delete(excel.styles)
            :ets.delete(excel.shared_strings)
            result
          {:error, reason} -> {:error, reason}
        end
      false ->
        excel = if timer do
          {_, s, ms} = :erlang.timestamp
          %{excel | timer: [s, ms]}
        else
          excel
        end

        case Unzip.validate_path_and_index(path, index) do
          {:ok, file}      -> extract_xml(file, index)
                              |> case do
                                {:ok, file_paths} ->
                                  {excel, result} = do_extract(excel, file_paths, index)
                                  if initial_parse do
                                    :ets.delete(excel.styles)
                                    :ets.delete(excel.shared_strings)
                                    result
                                  else
                                    {excel, result}
                                  end
                                {:error, reason}  -> {:error, reason}
                              end
          {:error, reason} -> {:error, reason}
        end
    end
  end

  @doc """
  Accesses ETS process and returns data formatted as a list of row value lists.

  ## Parameters
  - `tid` - table identifier of ETS process to be accessed

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> {:ok, tid} = Xlsxir.extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_list(tid)
          [["string one", "string two", 10, 20, {2016, 1, 1}]]
          iex> Xlsxir.close(tid)
          :ok

          iex> {:ok, tid} = Xlsxir.multi_extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_list(tid)
          [["string one", "string two", 10, 20, {2016, 1, 1}]]
          iex> Xlsxir.close(tid)
          :ok
  """
  def get_list(tid) do
    :ets.match(tid, {:"$1", :"$2"})
    |> Enum.sort
    |> Enum.map(fn [_num, row] ->
      row
      |> do_get_row()
      |> Enum.map(fn [_ref, val] -> val end)
    end)
  end

  @doc """
  Accesses ETS process and returns data formatted as a map of cell references and values.

  ## Parameters
  - `tid` - table identifier of ETS process to be accessed

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> {:ok, tid} = Xlsxir.extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_map(tid)
          %{ "A1" => "string one", "B1" => "string two", "C1" => 10, "D1" => 20, "E1" => {2016,1,1}}
          iex> Xlsxir.close(tid)
          :ok

          iex> {:ok, tid} = Xlsxir.multi_extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_map(tid)
          %{ "A1" => "string one", "B1" => "string two", "C1" => 10, "D1" => 20, "E1" => {2016,1,1}}
          iex> Xlsxir.close(tid)
          :ok
  """
  def get_map(tid) do
    :ets.match(tid, {:"$1", :"$2"})
    |> Enum.reduce(%{}, fn [_num, row], acc ->
         row |> do_get_row()
         |> Enum.reduce(%{}, fn [ref, val], acc2 -> Map.put(acc2, ref, val) end)
         |> Enum.into(acc)
       end)
  end

  @doc """
  Accesses ETS process and returns an indexed map which functions like a multi-dimensional array in other languages.

  ## Parameters
  - `tid` - table identifier of ETS process to be accessed

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> {:ok, tid} = Xlsxir.extract("./test/test_data/test.xlsx", 0)
          iex> mda = Xlsxir.get_mda(tid)
          %{0 => %{0 => "string one", 1 => "string two", 2 => 10, 3 => 20, 4 => {2016,1,1}}}
          iex> mda[0][0]
          "string one"
          iex> mda[0][2]
          10
          iex> Xlsxir.close(tid)
          :ok

          iex> {:ok, tid} = Xlsxir.multi_extract("./test/test_data/test.xlsx", 0)
          iex> mda = Xlsxir.get_mda(tid)
          %{0 => %{0 => "string one", 1 => "string two", 2 => 10, 3 => 20, 4 => {2016,1,1}}}
          iex> mda[0][0]
          "string one"
          iex> mda[0][2]
          10
          iex> Xlsxir.close(tid)
          :ok
  """
  def get_mda(tid) do
    :ets.match(tid, {:"$1", :"$2"}) |> convert_to_indexed_map(%{})
  end

  defp convert_to_indexed_map([], map), do: map

  defp convert_to_indexed_map([h|t], map) do
    row_index = Enum.at(h, 0)
                |> Kernel.-(1)

    add_row   = Enum.at(h,1)
                |> do_get_row()
                |> Enum.reduce({%{}, 0}, fn cell, {acc, index} ->
                  {Map.put(acc, index, Enum.at(cell, 1)), index + 1}
                end)
                |> elem(0)

    updated_map = Map.put(map, row_index, add_row)
    convert_to_indexed_map(t, updated_map)
  end

  @doc """
  Accesses ETS process and returns value of specified cell.

  ## Parameters
  - `table_id` - table identifier of ETS process to be accessed
  - `cell_ref` - Reference name of cell to be returned in `string` format (i.e. `"A1"`)

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> {:ok, tid} = Xlsxir.extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_cell(tid, "A1")
          "string one"
          iex> Xlsxir.close(tid)
          :ok

          iex> {:ok, tid} = Xlsxir.multi_extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_cell(tid, "A1")
          "string one"
          iex> Xlsxir.close(tid)
          :ok
  """
  def get_cell(table_id, cell_ref), do: do_get_cell(cell_ref, table_id)

  defp do_get_cell(cell_ref, table_id) do
    [[row_num]] = ~r/\d+/ |> Regex.scan(cell_ref)
    row_num     = row_num |> String.to_integer
    [[row]]     = :ets.match(table_id, {row_num, :"$1"})

    row
    |> Enum.filter(fn [ref, _val] -> ref == cell_ref end)
    |> List.first
    |> Enum.at(1)
  end

  @doc """
  Accesses ETS process and returns values of specified row in a `list`.

  ## Parameters
  - `tid` - table identifier of ETS process to be accessed
  - `row` - Reference name of row to be returned in `integer` format (i.e. `1`)

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> {:ok, tid} = Xlsxir.extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_row(tid, 1)
          ["string one", "string two", 10, 20, {2016, 1, 1}]
          iex> Xlsxir.close(tid)
          :ok

          iex> {:ok, tid} = Xlsxir.multi_extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_row(tid, 1)
          ["string one", "string two", 10, 20, {2016, 1, 1}]
          iex> Xlsxir.close(tid)
          :ok
  """
  def get_row(tid, row) do
    case :ets.match(tid, {row, :"$1"}) do
      [[row]] -> row |> do_get_row() |> Enum.map(fn [_ref, val] -> val end)
      [] -> []
    end
  end

  defp do_get_row(row) do
    row |> Enum.reduce({[], nil}, fn [ref, val], {values, previous} ->
      line = Regex.run(~r/\d+$/, ref) |> List.first
      empty_cells = cond do
        is_nil(previous) && String.first(ref) != "A" -> fill_empty_cells("A#{line}", ref, line, [])
        !is_nil(previous) && !is_next_col(ref, previous) -> fill_empty_cells(previous, ref, line, [])
        true -> []
      end
      {values ++ empty_cells ++ [[ref, val]], ref}
    end)
    |> elem(0)
  end

  defp is_next_col(current, previous) do
    String.to_charlist(current) == next_col(previous)
  end

  defp next_col(ref) do
    chars = Regex.run(~r/^[A-Z]+/, ref) |> List.first |> String.to_charlist
    case List.last(chars) do
      # The last letter is "Z"
      90 -> (String.duplicate("A", Enum.count(chars)) |> String.to_charlist) ++ [65]
      char -> Enum.take(chars, Enum.count(chars) - 1) ++ [char + 1]
    end
  end

  defp fill_empty_cells(from, from, _line, cells), do: Enum.reverse(cells)

  defp fill_empty_cells(from, to, line, cells) do
    next_ref = "#{next_col(from)}#{line}"
    if next_ref == to do
      fill_empty_cells(to, to, line, cells)
    else
      fill_empty_cells(next_ref, to, line, [[from, nil] | cells])
    end
  end

  @doc """
  Accesses `tid` ETS process and returns values of specified column in a `list`.

  ## Parameters
  - `tid` - table identifier of ETS process to be accessed
  - `col` - Reference name of column to be returned in `string` format (i.e. `"A"`)

  ## Example
  An example file named `test.xlsx` located in `./test/test_data` containing the following:
  - cell 'A1' -> "string one"
  - cell 'B1' -> "string two"
  - cell 'C1' -> integer of 10
  - cell 'D1' -> formula of "4 * 5"
  - cell 'E1' -> date of 1/1/2016 or Excel date serial of 42370

          iex> {:ok, tid} = Xlsxir.extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_col(tid, "A")
          ["string one"]
          iex> Xlsxir.close(tid)
          :ok

          iex> {:ok, tid} = Xlsxir.multi_extract("./test/test_data/test.xlsx", 0)
          iex> Xlsxir.get_col(tid, "A")
          ["string one"]
          iex> Xlsxir.close(tid)
          :ok
  """
  def get_col(tid, col), do: do_get_col(col, tid)

  defp do_get_col(col, tid) do
    :ets.match(tid, {:"$1", :"$2"})
    |> Enum.sort
    |> Enum.map(fn [_num, row] ->
      row |> do_get_row() |> Enum.filter_map(
        fn [ref, _val] -> Regex.scan(~r/[A-Z]+/i, ref) == [[col]] end,
        fn [_ref, val] -> val end
      )
    end)
    |> List.flatten
  end

  @doc """
  See `get_multi_info\2` documentation.
  """
  def get_info(table_id, num_type \\ :all) do
    get_multi_info(table_id, num_type)
  end

  @doc """
  Returns count data based on `num_type` specified:
  - `:rows` - Returns number of rows contained in worksheet
  - `:cols` - Returns number of columns contained in worksheet
  - `:cells` - Returns number of cells contained in worksheet
  - `:all` - Returns a keyword list containing all of the above

  ## Parameters
  - `tid` - table identifier of ETS process to be accessed
  - `num_type` - type of count data to be returned (see above), defaults to `:all`
  """
  def get_multi_info(tid, num_type \\ :all) do
    case num_type do
    :rows  -> row_num(tid)
    :cols  -> col_num(tid)
    :cells -> cell_num(tid)
    _      -> [
                rows:  row_num(tid),
                cols:  col_num(tid),
                cells: cell_num(tid)
              ]
    end
  end

  defp row_num(tid) do
    :ets.info(tid, :size)
  end

  defp col_num(tid) do
    :ets.match(tid, {:"$1", :"$2"})
    |> Enum.map(fn [_num, row] -> Enum.count(row) end)
    |> Enum.max
  end

  defp cell_num(tid) do
    :ets.match(tid, {:"$1", :"$2"})
    |> Enum.reduce(0, fn [_num, row], acc -> acc + Enum.count(row) end)
  end

  @doc """
  Deletes ETS process `tid` and returns `:ok` if successful.

  ## Example
  Extract first worksheet in an example file named `test.xlsx` located in `./test/test_data`:

      iex> {:ok, tid} = Xlsxir.extract("./test/test_data/test.xlsx", 0)
      iex> Xlsxir.close(tid)
      :ok

      iex> {:ok, tid} = Xlsxir.multi_extract("./test/test_data/test.xlsx", 0)
      iex> Xlsxir.close(tid)
      :ok
  """
  def close(tid) do
    if Enum.member?(:ets.all, tid) do
      if :ets.delete(tid), do: :ok, else: :error
    else
      :ok
    end
  end

  defp stop_timer(timer) do
    {_, s, ms} = :erlang.timestamp

    seconds      = s  |> Kernel.-(timer |> Enum.at(0))
    microseconds = ms |> Kernel.+(timer |> Enum.at(1))

    [add_s, micro] = if microseconds > 1_000_000 do
                       [1, microseconds - 1_000_000]
                     else
                       [0, microseconds]
                     end

    [h, m, s] = [
                  seconds / 3600 |> Float.floor |> round,
                  rem(seconds, 3600) / 60 |> Float.floor |> round,
                  rem(seconds, 60)
                ]

    [h, m, s + add_s, micro]
  end
end
