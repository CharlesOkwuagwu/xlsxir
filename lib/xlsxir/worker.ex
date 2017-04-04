defmodule Xlsxir.Worker do
  use GenServer

  def start_link() do
    GenServer.start_link(__MODULE__, %{"worksheet" => 0})
  end

  def handle_call({:table_id, id, value}, state), do: {:reply, :ok, Map.put(state, id, value)}

  def handle_call(:table_id, _from, state), do: {:reply, :ok, state}

  def handle_call(:codepoint, _from, state), do: {:reply, :ok, Map.put(state, "codepoint", 0)}

  def handle_call({:codepoint, codepoint}, _from, state), do: {:reply, :ok, Map.put(state, "codepoint", codepoint)}

  def handle_call(:get_codepoint, _from, state), do: {:reply, Map.get(state, "codepoint"), state}

  def handle_call(:rm_codepoint, _from, state), do: {:reply, :ok, Map.delete(state, "codepoint")}

  def handle_call(:worksheet, _from, state) do
    table = :ets.new(:worksheet, [:set, :protected])
    {:reply, table, Map.put(state, "worksheet", table)}
  end

  def handle_call(:get_worksheet, _from, %{"worksheet" => pid} = state), do: {:reply, pid, state}

  def handle_call({:worksheet, row, index}, _from, %{"worksheet" => pid} = state) do
    :ets.insert(pid, {index, row})
    {:reply, :ok, state}
  end

  def handle_call({:worksheet, row_num}, _from, %{"worksheet" => pid} = state) do
    row = :ets.lookup(pid, row_num)
    row = if row == [] do
      row
    else
      row
      |> List.first
      |> Tuple.to_list
      |> Enum.at(1)
    end
    {:reply, row, state}
  end

  def handle_call(:shared_strings, _from, state) do
    table = :ets.new(:shared_string, [:set, :protected])
    {:reply, table, Map.put(state, "shared_strings", table)}
  end

  def handle_call({:shared_strings, string, index}, _from, %{"shared_strings" => pid} = state) do
    :ets.insert(pid, {index, string})
    {:reply, :ok, state}
  end

  def handle_call(:rm_shared_strings, _from, state) do
    case Map.get(state, "shared_strings") do
      nil -> nil
      pid -> :ets.delete(pid)
    end
    {:reply, :ok, Map.delete(state, "shared_strings")}
  end

  def handle_call({:shared_strings, index}, _from, %{"shared_strings" => pid} = state) do
    value = :ets.lookup(pid, index)
    |> List.first
    |> Tuple.to_list
    |> Enum.at(1)
    {:reply, value, state}
  end

  def handle_call(:styles, _from, state) do
    table = :ets.new(:styles, [:set, :protected])
    {:reply, table, Map.put(state, "styles", table) |> Map.put("num_fmt_ids", [])}
  end

  def handle_call({:add_styles_id, num_fmt_id}, _from, %{"num_fmt_ids" => num_fmt_ids} = state) do
    {:reply, :ok, Map.put(state, "num_fmt_ids", Enum.into([num_fmt_id], num_fmt_ids))}
  end

  def handle_call(:rm_styles_id, _from, state) do
    {:reply, :ok, Map.delete(state, "num_fmt_ids")}
  end

  def handle_call(:get_style_ids, _from, %{"num_fmt_ids" => num_fmt_ids} = state) do
    {:reply, num_fmt_ids, state}
  end

  def handle_call({:styles, style, index}, _from, %{"styles" => pid} = state) do
    {:reply, :ets.insert(pid, {index, style}), state}
  end

  def handle_call({:styles, index}, _from, %{"styles" => pid} = state) do
    value = :ets.lookup(pid, index)
    |> List.first
    |> Tuple.to_list
    |> Enum.at(1)
    {:reply, value, state}
  end

  def handle_call(:rm_styles, _from, state) do
    case Map.get(state, "styles") do
      nil -> nil
      pid -> :ets.delete(pid)
    end
    {:reply, :ok, Map.delete(state, "styles")}
  end

  def handle_call(:index, _from, state) do
    {:reply, :ok, Map.put(state, "index", 0)}
  end

  def handle_call(:increment_1, _from, %{"index" => index} = state) do
    {:reply, :ok, Map.put(state, "index", index + 1)}
  end

  def handle_call(:get_index, _from, state) do
    {:reply, Map.get(state, "index"), state}
  end

  def handle_call(:rm_index, _from, state) do
    {:reply, :ok, Map.delete(state, "index")}
  end

  def handle_call(:timer, _from, state) do
    {_, s, ms} = :erlang.timestamp
    {:reply, :ok, Map.put(state, "timer", [s, ms])}
  end

  def handle_call(:stop_timer, _from, %{"timer" => timer} = state) do
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

    {:reply, [h, m, s + add_s, micro], Map.delete(state, "timer")}
  end

  def handle_call({:max_rows, max_rows}, _from, state) do
    {:reply, :ok, Map.put(state, "max_rows", max_rows)}
  end

  def handle_call(:get_max_rows, _from, state) do
    {:reply, Map.get(state, "max_rows"), state}
  end

  def handle_call(:rm_max_rows, _from, state) do
    {:reply, :ok, Map.delete(state, "max_rows")}
  end
end
