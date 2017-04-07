defmodule Xlsxir.Mixfile do
  use Mix.Project

  def project do
    [
     app: :xlsxir,
     version: "1.5.1",
     name: "Xlsxir",
     source_url: "https://github.com/kennellroxco/xlsxir",
     elixir: "~> 1.4",
     build_embedded: Mix.env == :prod,
     start_permanent: Mix.env == :prod,
     description: description(),
     package: package(),
     deps: deps(),
     docs: [main: "overview", extras: ["CHANGELOG.md", "NUMBER_STYLES.md", "OVERVIEW.md"]]
    ]
  end

  def application do
    [applications: [:logger], mod: {Xlsxir, []}]
  end

  defp deps do
    [
      { :ex_doc, github: "elixir-lang/ex_doc", only: :dev },
      { :earmark, github: "pragdave/earmark", override: true, only: :dev },
      { :erlsom, "~> 1.4" }
    ]
  end

  defp description do
    """
    Xlsx file parser
    """
  end

  defp package do
    [
      maintainers: ["Jason Kennell", "Bryan Weatherly"],
      licenses: ["MIT License"],
      links: %{
                "Github" => "https://github.com/kennellroxco/xlsxir",
                "Change Log" => "https://hexdocs.pm/xlsxir/changelog.html"
               }
    ]
  end

end
