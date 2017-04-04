defmodule Xlsxir.ParseWorksheet do
  alias Xlsxir.{ConvertDate, ConvertDateTime, SaxError}
  import Xlsxir.ConvertDate, only: [convert_char_number: 1]
  require Logger
  @moduledoc """
  Holds the SAX event instructions for parsing worksheet data via `Xlsxir.SaxParser.parse/2`
  """

  defstruct row: %{}, cell_ref: "", data_type: "", num_style: "", value: "", max_rows: nil

  @doc """
  Sax event utilized by `Xlsxir.SaxParser.parse/2`. Takes a pattern and the current state of a struct and recursivly parses the
  worksheet XML file, ultimately sending a keyword list of cell references and their assocated values to the `Xlsxir.Worksheet` module
  which contains an ETS process that was started by `Xlsxir.SaxParser.parse/2`.

  ## Parameters

  - `arg1` - the XML pattern of the event to match upon
  - `state` - the state of the `%Xlsxir.ParseWorksheet{}` struct which temporarily holds applicable data of the current row being parsed

  ## Example
  Each entry in the list created consists of a list containing a cell reference string and the associated value (i.e. `[["A1", "string one"], ...]`).
  """
  def sax_event_handler({:startElement,_,'row',_,_}, _state, pid) do
    GenServer.call(pid, :codepoint)
    %Xlsxir.ParseWorksheet{
      max_rows: GenServer.call(pid, :get_max_rows)
    }
  end

  def sax_event_handler({:startElement,_,'c',_,xml_attr}, state, pid) do
    a = Enum.map(xml_attr, fn(attr) ->
          case attr do
            {:attribute,'r',_,_,ref}   -> {:r, ref  }
            {:attribute,'s',_,_,style} -> {:s, GenServer.call(pid, {:styles, List.to_integer(style)})}
            {:attribute,'t',_,_,type}  -> {:t, type }
            _                          -> raise "Unknown cell attribute"
          end
        end)

    {cell_ref, num_style, data_type} = case Keyword.keys(a) |> Enum.sort do
                                         [:r]         -> {a[:r],   nil,   nil}
                                         [:r, :s]     -> {a[:r], a[:s],   nil}
                                         [:r, :t]     -> {a[:r],   nil, a[:t]}
                                         [:r, :s, :t] -> {a[:r], a[:s], a[:t]}
                                         _            -> raise "Invalid attributes: #{a}"
                                       end
    %{state | cell_ref: cell_ref, num_style: num_style, data_type: data_type}
  end

  def sax_event_handler({:characters, value}, state, _) do
    if state == nil, do: nil, else: %{state | value: value}
  end

  def sax_event_handler({:endElement,_,'c',_}, %Xlsxir.ParseWorksheet{row: row} = state, pid) do
    cell_value = format_cell_value(pid, [state.data_type, state.num_style, state.value])
    new_cells  = compare_to_previous_cell(pid, to_string(state.cell_ref), cell_value)
    %{state | row: Enum.into(row, new_cells), cell_ref: "", data_type: "", num_style: "", value: ""}
  end

  def sax_event_handler({:endElement,_,'row',_}, state, pid) do
    GenServer.call(pid, :rm_codepoint)

    unless Enum.empty?(state.row) do
      [[row]] = ~r/\d+/ |> Regex.scan(state.row |> List.first |> List.first)
      row     = row |> String.to_integer
      value = state.row |> Enum.reverse
      GenServer.call(pid, {:worksheet, value, row})
      if !is_nil(state.max_rows) and row == state.max_rows, do: raise SaxError
    end
    state
  end

  def sax_event_handler(:endDocument, _state, pid) do
    GenServer.call(pid, :rm_shared_strings)
    GenServer.call(pid, :rm_styles)
  end

  def sax_event_handler(_, state, _), do: state

  defp format_cell_value(pid, list) do
    case list do
      [          _,   _, nil] -> nil                                                                 # Cell with no value attribute
      [          _,   _,  ""] -> nil                                                                 # Empty cell with assigned attribute
      [        'e', nil,   e] -> List.to_string(e)                                                   # Type error
      [        's',   _,   i] -> GenServer.call(pid, {:shared_strings, List.to_integer(i)})           # Type string
      [        nil, nil,   n] -> convert_char_number(n)                                              # Type number
      [        'n', nil,   n] -> convert_char_number(n)
      [        nil, 'd',   d] -> convert_date_or_time(d)                                             # ISO 8601 type date
      [        'n', 'd',   d] -> convert_date_or_time(d)
      [      'str',   _,   s] -> List.to_string(s)                                                   # Type formula w/ string
      [        'b',   _,   s] -> s == '1'                                                            # Type boolean
      ['inlineStr',   _,   s] -> List.to_string(s)                                                   # Type string
      _                       -> raise "Unmapped attribute #{Enum.at(list, 0)}. Unable to process"   # Unmapped type
    end
  end

  defp convert_date_or_time(value) do
    str = List.to_string(value)

    if str == "0" || String.match?(str, ~r/\d\.\d+/) do
      ConvertDateTime.from_charlist(value)
    else
      ConvertDate.from_serial(value)
    end
  end

  defp compare_to_previous_cell(pid, ref, value) do
    [[<<codepoint::utf8>>]] = ~r/[A-Z](?=[0-9])/i |> Regex.scan(ref)
    [[row_num]]             = ~r/[0-9]+/          |> Regex.scan(ref)
    [[col_ltr]]             = ~r/[A-Z]+/i         |> Regex.scan(ref)

    prefix = col_ltr |> String.slice(0, String.length(col_ltr) - 1)
    current_codepoint = GenServer.call(pid, :get_codepoint)

    cond do
      current_codepoint == 0 and codepoint == 65  -> GenServer.call(pid, {:codepoint, 65})
                                                      [[ref, value]]
      current_codepoint == 90 and codepoint == 65 -> GenServer.call(pid, {:codepoint, 65})
                                                      [[ref, value]]
      codepoint - current_codepoint == 1          -> GenServer.call(pid, {:codepoint, codepoint})
                                                      [[ref, value]]
      true                                        -> get_empty_cells(pid, codepoint, row_num, prefix)
                                                      |> Enum.into([[ref, value]])
    end
  end

  defp get_empty_cells(pid, codepoint, row_num, prefix) do
    current_codepoint = GenServer.call(pid, :get_codepoint)
    range = cond do
      current_codepoint == 0        -> (codepoint - 1)..65
      codepoint == 65               -> 90..(current_codepoint + 1)
      current_codepoint == 90       -> (codepoint - 1)..65
      current_codepoint > codepoint -> [90..(current_codepoint + 1), (codepoint - 1)..65]
      true                          -> (codepoint - 1)..(current_codepoint + 1)
    end

    GenServer.call(pid, {:codepoint, codepoint})
    create_empty_cells(range, row_num, prefix)
  end

  defp create_empty_cells(range, row_num, prefix) when is_list(range) do
    [third, second] = case String.split(prefix, "", trim: true) do
                        [<<second::utf8>>]                  -> ["", second]
                        [<<third::utf8>>, <<second::utf8>>] -> [third, second]
                      end

    [pre_z, post_z] = range

    case String.length(prefix) do
      1 -> case second do
             65 -> tail = pre_z  |> Enum.map(fn col_ltr -> [              <<col_ltr>> <> row_num, nil] end)
                   head = post_z |> Enum.map(fn col_ltr -> [<<second>> <> <<col_ltr>> <> row_num, nil] end)
                   Enum.into(tail, head)

              _ -> tail = pre_z  |> Enum.map(fn col_ltr -> [<<(second - 1)>> <> <<col_ltr>> <> row_num, nil] end)
                   head = post_z |> Enum.map(fn col_ltr -> [      <<second>> <> <<col_ltr>> <> row_num, nil] end)
                   Enum.into(tail, head)
           end

      2 -> case [third, second] do
             [65, 65] -> tail = pre_z  |> Enum.map(fn col_ltr -> [                 <<90, col_ltr>> <> row_num, nil] end)
                         head = post_z |> Enum.map(fn col_ltr -> [<<third, second>> <> <<col_ltr>> <> row_num, nil] end)
                         Enum.into(tail, head)

                    _ -> tail = pre_z  |> Enum.map(fn col_ltr -> [<<third, (second - 1)>> <> <<col_ltr>> <> row_num, nil] end)
                         head = post_z |> Enum.map(fn col_ltr -> [      <<third, second>> <> <<col_ltr>> <> row_num, nil] end)
                         Enum.into(tail, head)
           end
    end
  end

  defp create_empty_cells(range, row_num, prefix) do
    Enum.map(range, fn col_ltr -> [prefix <> <<col_ltr>> <> row_num, nil] end)
  end

end
