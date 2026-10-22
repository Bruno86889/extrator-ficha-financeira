
from dataclasses import dataclass


@dataclass(frozen=True)
class Lancamento:
    ano: int | None
    mes: str | None
    denominacao: str
    vd: str
    op: str
    nat: str
    qtde: str | None
    periodo_inicio: str | None
    periodo_fim: str | None
    secao: str
    pagina: int
    linha_original: str


@dataclass(frozen=True)
class ResultadoExtracao:
    lancamentos: tuple[Lancamento, ...]
    denominacaoes: tuple[str, ...]
