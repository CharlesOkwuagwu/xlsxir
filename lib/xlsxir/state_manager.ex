defmodule Xlsxir.StateManager do
  use GenServer

  def start_link() do
    GenServer.start_link(__MODULE__, :ok, name: __MODULE__)
  end

  def handle_call(:new_table, _from, state) do
    {:reply, :ets.new(:table, [:set, :public]), state}
  end
end
