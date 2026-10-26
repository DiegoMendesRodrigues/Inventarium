# Sistema criado por Diego Mendes Rodrigues, utilizando desenvolvimento de software e inteligência artificial
"""Envia avisos de validade de licenças de software aos administradores inscritos."""

from datetime import date, timedelta
from pathlib import Path
import sys


RAIZ_PROJETO = Path(__file__).resolve().parents[1]
if str(RAIZ_PROJETO) not in sys.path:
    sys.path.insert(0, str(RAIZ_PROJETO))

from app import Administrador, Software, app, enviar_email_softwares_assinatura, softwares_que_irao_expirar


def main():
    with app.app_context():
        hoje = date.today()
        limite = hoje + timedelta(days=60)
        quantidade = Software.query.filter(
            Software.validade_assinatura.isnot(None),
            Software.validade_assinatura >= hoje,
            Software.validade_assinatura <= limite,
        ).count()

        if quantidade == 0:
            print('Nenhum software com assinatura expirando nos próximos 60 dias.')
            return 0

        lista_softwares = softwares_que_irao_expirar()
        administradores = Administrador.query.filter_by(recebe_notificacoes='Sim') \
            .order_by(Administrador.id.asc()).all()
        if not administradores:
            print('Nenhum administrador está inscrito para receber notificações do sistema.')
            return 0

        falhas = 0
        for administrador in administradores:
            try:
                enviar_email_softwares_assinatura(administrador, lista_softwares)
            except Exception as erro:
                falhas += 1
                print(f'Falha ao enviar e-mail para {administrador.nome_completo}: {erro}')
            else:
                print(f'E-mail enviado para {administrador.nome_completo}.')

        return 1 if falhas else 0


if __name__ == '__main__':
    raise SystemExit(main())
