defmodule ElixirExcel do
  @moduledoc """
  Documentation for ElixirExcel.
  """

  @doc """
  Hello world.

  ## Examples

      iex> ElixirExcel.hello
      :world

  """
  def hello do
    :world
  end

  def parse(file_path) do
    file_path = file_path |> Path.expand()
    extname = file_path |> Path.extname()
    _parse(file_path, extname)
  end

  defp _parse(file_path, extname) when extname in ~W(.xlsx .xls) do
    cmd_file =
      :code.priv_dir(:elixir_excel)
      |> Path.join("nodejs")
      |> Path.join("parseExcel")

    case System.cmd("node", [cmd_file, "-f", file_path]) do
      {"ERROR:MESSAGE:" <> message, 0} ->
        {:error, :parse_fail, String.trim(message)}

      {"OK:EXCEL:" <> excel, 0} ->
        case Jason.decode(excel, keys: :strings) do
          {:ok, value} ->
            {:ok, value}

          {:error, reason} ->
            {:error, reason, "#{reason}"}

          _ ->
            {:error, :parse_fail, "parse fail"}
        end

      _ ->
        {:error, :parse_fail, "parse fail"}
    end
  end

  defp _parse(_file_path, extname) do
    {:error, :nonsupport, "ext #{extname} unsupported"}
  end
end
