
from datetime import datetime
from sqlalchemy import (
    Column,
    Integer,
    String,
    Date,
    DateTime,
    ForeignKey,
    Text,
)
from sqlalchemy.orm import relationship

from database import Base


class Cliente(Base):
    __tablename__ = "clientes"

    id = Column(Integer, primary_key=True)
    organization_id = Column(Integer, ForeignKey("organizations.id"), nullable=True)
    razao_social = Column(String, nullable=False)
    contato = Column(String, nullable=True)
    cnpj = Column(String, nullable=True, unique=False)
    ie = Column(String, nullable=True)

    rua = Column(String, nullable=True)
    numero = Column(String, nullable=True)
    bairro = Column(String, nullable=True)
    cidade = Column(String, nullable=True)
    uf = Column(String, nullable=True)
    cep = Column(String, nullable=True)

    ddd = Column(String, nullable=True)
    telefone = Column(String, nullable=True)
    email = Column(String, nullable=True)

    ultima_inspecao = Column(Date, nullable=True)

    organization = relationship("Organization", back_populates="clientes")
    entradas_relatorio = relationship("EntradaRelatorio", back_populates="cliente")

    def __repr__(self) -> str:
        return f"<Cliente(id={self.id}, razao_social='{self.razao_social}')>"


class TipoRelatorio(Base):
    __tablename__ = "tipos_relatorio"

    id = Column(Integer, primary_key=True)
    nome = Column(String, nullable=False, unique=True)
    descricao = Column(String, nullable=True)
    schema_json = Column(Text, nullable=False)
    template_path = Column(String, nullable=True)

    entradas = relationship("EntradaRelatorio", back_populates="tipo_relatorio")

    def __repr__(self) -> str:
        return f"<TipoRelatorio(id={self.id}, nome='{self.nome}')>"


class EntradaRelatorio(Base):
    __tablename__ = "entradas_relatorio"

    id = Column(Integer, primary_key=True)
    organization_id = Column(Integer, ForeignKey("organizations.id"), nullable=True)
    cliente_id = Column(Integer, ForeignKey("clientes.id"), nullable=True)
    tipo_relatorio_id = Column(Integer, ForeignKey("tipos_relatorio.id"), nullable=False)

    relatorio_num = Column(String, nullable=True)
    titulo_personalizado = Column(String, nullable=True)
    dados_json = Column(Text, nullable=True)
    caminho_arquivo_gerado = Column(String, nullable=True)
    criado_em = Column(DateTime, default=datetime.now)

    organization = relationship("Organization", back_populates="entradas_relatorio")
    cliente = relationship("Cliente", back_populates="entradas_relatorio")
    tipo_relatorio = relationship("TipoRelatorio", back_populates="entradas")

    def __repr__(self) -> str:
        return f"<EntradaRelatorio(id={self.id}, num='{self.relatorio_num}', tipo_id={self.tipo_relatorio_id})>"


class Insumo(Base):
    __tablename__ = "insumos"

    id = Column(Integer, primary_key=True)
    organization_id = Column(Integer, ForeignKey("organizations.id"), nullable=False)
    tipo = Column(String, nullable=False)
    nome = Column(String, nullable=False)
    fabricante = Column(String, nullable=True)
    data_fabricacao = Column(String, nullable=True)
    data_validade = Column(String, nullable=True)
    lote = Column(String, nullable=True)
    criado_em = Column(DateTime, default=datetime.now)

    organization = relationship("Organization", back_populates="insumos")

    def __repr__(self) -> str:
        return f"<Insumo(id={self.id}, tipo='{self.tipo}', nome='{self.nome}')>"


class Organization(Base):
    __tablename__ = "organizations"

    id = Column(Integer, primary_key=True)
    nome = Column(String, nullable=False)
    criado_em = Column(DateTime, default=datetime.now)

    users = relationship("User", back_populates="organization")
    clientes = relationship("Cliente", back_populates="organization")
    entradas_relatorio = relationship("EntradaRelatorio", back_populates="organization")
    insumos = relationship("Insumo", back_populates="organization")

    def __repr__(self) -> str:
        return f"<Organization(id={self.id}, nome='{self.nome}')>"


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True)
    organization_id = Column(Integer, ForeignKey("organizations.id"), nullable=False)
    nome = Column(String, nullable=False)
    email = Column(String, nullable=False, unique=True)
    password_hash = Column(String, nullable=False)
    role = Column(String, nullable=False, default="admin")
    criado_em = Column(DateTime, default=datetime.now)

    organization = relationship("Organization", back_populates="users")

    def __repr__(self) -> str:
        return f"<User(id={self.id}, email='{self.email}')>"
