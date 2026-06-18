import os
from datetime import datetime

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, EmailStr, Field
from sqlalchemy import DateTime, Integer, String, Text, create_engine
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column


DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./app_users.db")


def normalize_database_url(url: str) -> str:
    # Railway commonly provides postgres:// or postgresql:// URLs.
    # Force psycopg driver explicitly so SQLAlchemy does not try psycopg2.
    if url.startswith("postgres://"):
        return "postgresql+psycopg://" + url[len("postgres://") :]
    if url.startswith("postgresql://") and "+" not in url.split("://", 1)[0]:
        return "postgresql+psycopg://" + url[len("postgresql://") :]
    return url


def build_engine(database_url: str):
    normalized = normalize_database_url(database_url)
    if normalized.startswith("sqlite"):
        return create_engine(
            normalized,
            echo=False,
            future=True,
            connect_args={"check_same_thread": False},
        )
    return create_engine(
        normalized,
        echo=False,
        future=True,
        pool_pre_ping=True,
    )


class Base(DeclarativeBase):
    pass


class AppUser(Base):
    __tablename__ = "app_users"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(120), nullable=False)
    email: Mapped[str] = mapped_column(String(255), unique=True, nullable=False)
    role: Mapped[str] = mapped_column(String(50), default="user", nullable=False)
    datasheets_root: Mapped[str] = mapped_column(Text, nullable=False)
    sharepoint_library_url: Mapped[str] = mapped_column(Text, nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, nullable=False)


engine = build_engine(DATABASE_URL)

Base.metadata.create_all(engine)

app = FastAPI(title="Datasheets User DB API", version="0.1.0")


class UserCreate(BaseModel):
    name: str = Field(min_length=2, max_length=120)
    email: EmailStr
    role: str = Field(default="user", max_length=50)
    datasheets_root: str
    sharepoint_library_url: str


class UserOut(BaseModel):
    id: int
    name: str
    email: str
    role: str
    datasheets_root: str
    sharepoint_library_url: str
    created_at: datetime


@app.get("/health")
def health() -> dict:
    return {"status": "ok", "database_url_set": bool(DATABASE_URL)}


@app.post("/users", response_model=UserOut)
def create_user(payload: UserCreate) -> UserOut:
    with Session(engine) as session:
        existing = session.query(AppUser).filter(AppUser.email == str(payload.email)).first()
        if existing:
            raise HTTPException(status_code=409, detail="Email already exists")

        user = AppUser(
            name=payload.name,
            email=str(payload.email),
            role=payload.role,
            datasheets_root=payload.datasheets_root,
            sharepoint_library_url=payload.sharepoint_library_url,
        )
        session.add(user)
        session.commit()
        session.refresh(user)

        return UserOut(
            id=user.id,
            name=user.name,
            email=user.email,
            role=user.role,
            datasheets_root=user.datasheets_root,
            sharepoint_library_url=user.sharepoint_library_url,
            created_at=user.created_at,
        )


@app.get("/users", response_model=list[UserOut])
def list_users() -> list[UserOut]:
    with Session(engine) as session:
        rows = session.query(AppUser).order_by(AppUser.id.asc()).all()

        return [
            UserOut(
                id=row.id,
                name=row.name,
                email=row.email,
                role=row.role,
                datasheets_root=row.datasheets_root,
                sharepoint_library_url=row.sharepoint_library_url,
                created_at=row.created_at,
            )
            for row in rows
        ]


@app.delete("/users/{user_id}")
def delete_user(user_id: int) -> dict:
    with Session(engine) as session:
        row = session.get(AppUser, user_id)
        if not row:
            raise HTTPException(status_code=404, detail="User not found")

        session.delete(row)
        session.commit()

    return {"deleted": user_id}
