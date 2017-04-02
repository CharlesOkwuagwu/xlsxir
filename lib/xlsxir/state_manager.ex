defmodule Xlsxir.TableId do
  @moduledoc """
  An Agent process named `TableId` which temporarily holds the table id of an ETS process. Provides functions to create the process, assign a table id, retrieve the current table id
  and ultimately kill the process.
  """

  @doc """
  Initiates a new `TableId` Agent process with a value of `0`.
  """
  def new do
    Agent.start_link(fn -> %{} end)
  end

  @doc """
  Assigns an ETS table id to the Agent process.
  """
  def assign_id(pid, id, value) when is_pid(pid) do
    Agent.update(pid, fn state -> Map.put(state, id, value) end)
  end

  @doc """
  Returns current value of Agent process
  """
  def get(pid) when is_pid(pid)do
    Agent.get(pid, &(&1))
  end

  @doc """
  Deletes Agent process
  """
  def delete(pid) when is_pid(pid) do
    Agent.stop(pid)
  end

  def alive?(pid) when is_pid(pid) do
    Process.alive?(pid)
  end
end

defmodule Xlsxir.Codepoint do
  @moduledoc """
  An Agent process named `Codepoint` which temporarily holds the codepoint value of the last column letter of the most recently extracted cell. Provides functions to create the process,
  update the codepoint value being held, retrieve the currently held codepoint and ultimately kill the process.
  """

  @doc """
  Initiates a new `Codepoint` Agent process with a value of `0`.
  """
  def new do
    Agent.start_link(fn -> 0 end, name: Codepoint)
  end

  @doc """
  Updates the Agent process to the given codepoint
  """
  def hold(codepoint) do
    Agent.update(Codepoint, &(&1 - &1 + codepoint))
  end

  @doc """
  Returns current codepoing being held by the Agent process
  """
  def get do
    Agent.get(Codepoint, &(&1))
  end

  @doc """
  Deletes `Codepoint` Agent process
  """
  def delete do
    Agent.stop(Codepoint)
  end
end

defmodule Xlsxir.Worksheet do
  alias Xlsxir.TableId

  @moduledoc """
  An Erlang Term Storage (ETS) process named `:worksheet` which holds state for data parsed from `sheet\#{n}.xml` at index `n`. Provides functions to create the process, add
  and retreive data, and ultimately kill the process.
  """

  @doc """
  Initializes new ETS process with `[:set, :protected, :named_table]` options.
  """
  def new(pid) when is_pid(pid) do
    id = "#{pid |> inspect}_worksheet"
    table = :ets.new(id |> String.to_atom, [:set, :protected])
    TableId.assign_id(pid, id, table)
  end

  @doc """
  Initializes new ETS process with `[:set, :protected]` options and assigns the associated table id to `Xlsxir.TableId`.
  """

  #def new_multi(pid) do
    #id = "#{pid |> inspect}_multi_worksheet"
    #table = :ets.new(id |> String.to_atom, [:set, :protected])
    #TableId.assign_id(pid, id, table)
  #end

  @doc """
  Stores a row at a given index in the ETS process.
  """
  def add_row(row, index, pid) when is_pid(pid) do
    id_by_pid(pid) |> :ets.insert({index, row})
  end

  @doc """
  Returns a row at a given index of the ETS process.
  """
  def get_at(row_num, pid) when is_pid(pid) do
    row = :ets.lookup(id_by_pid(pid), row_num)

    if row == [] do
      row
    else
      row
      |> List.first
      |> Tuple.to_list
      |> Enum.at(1)
    end
  end

  @doc """
  Deletes the ETS process from memory.
  """
  def delete(pid) when is_pid(pid) do
    if alive?(pid), do: :ets.delete(id_by_pid(pid)), else: false
  end

  @doc """
  Validates whether the ETS process is active, returning true or false.
  """
  def alive?(pid) when is_pid(pid) do
    Enum.member?(:ets.all, id_by_pid(pid))
  end

  def id_by_pid(pid) when is_pid(pid) do
    "#{pid |> inspect}_worksheet" |> String.to_atom
  end
end

defmodule Xlsxir.SharedString do
  @moduledoc """
  An Erlang Term Storage (ETS) process named `:sharedstrings` which holds state for data parsed from `sharedStrings.xml`. Provides functions to create the process, add
  and retreive data, and ultimately kill the process.
  """

  alias Xlsxir.TableId

  @doc """
  Initializes new ETS process with `[:set, :protected, :named_table]` options.
  """
  def new(pid) when is_pid(pid) do
    id = id_by_pid(pid)
    table = :ets.new(id, [:set, :protected, :named_table])
    TableId.assign_id(pid, "sharedstring", table)
  end

  @doc """
  Stores a sharedstring at a given index in the ETS process.
  """
  def add_shared_string(pid, shared_string, index) when is_pid(pid) do
    id_by_pid(pid) |> :ets.insert({index, shared_string})
  end

  @doc """
  Returns a sharedstring at a given index of the ETS process.
  """
  def get_at(pid, index) when is_pid(pid) do
    :ets.lookup(id_by_pid(pid), index)
    |> List.first
    |> Tuple.to_list
    |> Enum.at(1)
  end

  @doc """
  Deletes the ETS process from memory.
  """
  def delete(pid) when is_pid(pid) do
    pid |> id_by_pid |> :ets.delete
  end

  @doc """
  Validates whether the ETS process is active, returning true or false.
  """
  def alive?(pid) when is_pid(pid) do
    Enum.member?(:ets.all, id_by_pid(pid))
  end

  def id_by_pid(pid) when is_pid(pid) do
    "#{pid |> inspect}_sharedstrings" |> String.to_atom
  end
end

defmodule Xlsxir.Style do
  @moduledoc """
  An Erlang Term Storage (ETS) process named `:styles` which holds state for data parsed from `styles.xml`. Provides functions to create the process, add and retreive data,
  and ultimately kill the process. Also includes a temporary Agent process named `NumFmtIds` which is utilized during the parsing of the `styles.xml` file to temporarily
  hold state of each `NumFmtId` contained within the file.
  """

  @doc """
  Initializes new ETS process with `[:set, :protected, :named_table]` options. Additionally, initiates an Agent process to temporarily hold `numFmtId`s for `Xlsxir.ParseStyle`.
  """
  def new do
    Agent.start_link(fn -> [] end, name: NumFmtIds)
    :ets.new(:styles, [:set, :protected, :named_table])
  end

  # functions for `NumFmtIds`
  @doc """
  Adds a `numFmtId` to the `NumFmtIds` Agent process.
  """
  def add_id(num_fmt_id) do
    Agent.update(NumFmtIds, &(Enum.into([num_fmt_id], &1)))
  end

  @doc """
  Returns a  list of `numFmtId`s stored in the `NumFmtIds` Agent process.
  """
  def get_id do
    Agent.get(NumFmtIds, &(&1))
  end

  @doc """
  Deletes `NumFmtIds` Agent process.
  """
  def delete_id do
    Agent.stop(NumFmtIds)
  end

  # functions for `:styles`
  @doc """
  Stores a style type at a given index in the ETS process.
  """
  def add_style(style, index) do
    :styles |> :ets.insert({index, style})
  end

  @doc """
  Returns a style type at a given index of the ETS process.
  """
  def get_at(index) do
    :ets.lookup(:styles, index)
    |> List.first
    |> Tuple.to_list
    |> Enum.at(1)
  end

  @doc """
  Deletes the ETS process from memory.
  """
  def delete do
    :ets.delete(:styles)
  end

  @doc """
  Validates whether the ETS process is active, returning true or false.
  """
  def alive? do
    Enum.member?(:ets.all, :styles)
  end
end

defmodule Xlsxir.Index do
  @moduledoc """
  An Agent process named `Index` which holds state of an index. Provides functions to create the process, increment the index by 1, retrieve the current index
  and ultimately kill the process.
  """

  @doc """
  Initiates a new `Index` Agent process with a value of `0`.
  """
  def new do
    Agent.start_link(fn -> 0 end, name: Index)
  end

  @doc """
  Increments active `Index` Agent process by `1`.
  """
  def increment_1 do
    Agent.update(Index, &(&1 + 1))
  end

  @doc """
  Returns current value of `Index` Agent process
  """
  def get do
    Agent.get(Index, &(&1))
  end

  @doc """
  Deletes `Index` Agent process
  """
  def delete do
    Agent.stop(Index)
  end
end

defmodule Xlsxir.Timer do
  @moduledoc """
  An `Agent` process named `Time` which holds state for time elapsed since execution. Provides functions to start and stop the process, with the stop function returning the time elapsed as a
  list (i.e. `[hour, minute, second, microsecond]`).
  """

  @doc """
  Initiates a new `Time` Agent process. Records current time via `:erlang.timestamp` and saves it to the Agent process.
  """
  def start do
    Agent.start_link(fn -> [] end, name: Time)
    {_, s, ms} = :erlang.timestamp
    Agent.update(Time, &(Enum.into([s, ms], &1)))
  end

  @doc """
  Records current time via `:erlang.timestamp` and compares it to the timestamp held by the `Time` Agent process to determin the amount of time elapsed. Returns the time elapsed in the format of
  `[hour, minute, second, microsecond]`.
  """
  def stop do
    {_, s, ms} = :erlang.timestamp

    seconds      = s  |> Kernel.-(Agent.get(Time, &(&1)) |> Enum.at(0))
    microseconds = ms |> Kernel.+(Agent.get(Time, &(&1)) |> Enum.at(1))

    [add_s, micro] = if microseconds > 1_000_000 do
                       [1, microseconds - 1_000_000]
                     else
                       [0, microseconds]
                     end

    [h, m, s] = [
                  seconds/3600 |> Float.floor |> round,
                  rem(seconds, 3600)/60 |> Float.floor |> round,
                  rem(seconds, 60)
                ]

    Agent.stop(Time)
    [h, m, s + add_s, micro]
  end
end
