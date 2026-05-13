from werkzeug.security import generate_password_hash

from database import get_session, init_db
from models import Organization, User


def main() -> None:
    init_db()
    email = input("E-mail do admin: ").strip().lower()
    nome = input("Nome do admin: ").strip() or "Administrador"
    empresa = input("Nome da empresa: ").strip() or "RL Metais"
    senha = input("Nova senha: ")

    if not email or len(senha) < 6:
        print("Informe um e-mail e uma senha com pelo menos 6 caracteres.")
        return

    session = get_session()
    try:
        user = session.query(User).filter_by(email=email).first()
        if user:
            user.nome = nome
            user.password_hash = generate_password_hash(senha)
            if user.organization:
                user.organization.nome = empresa
            print("Senha redefinida com sucesso.")
        else:
            org = session.query(Organization).first()
            if not org:
                org = Organization(nome=empresa)
                session.add(org)
                session.flush()
            else:
                org.nome = empresa

            user = User(
                organization_id=org.id,
                nome=nome,
                email=email,
                password_hash=generate_password_hash(senha),
                role="admin",
            )
            session.add(user)
            print("Usuário admin criado com sucesso.")

        session.commit()
    finally:
        session.close()


if __name__ == "__main__":
    main()
