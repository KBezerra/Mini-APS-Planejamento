<div align="center">

# ⚙️ Mini-APS — Sistema de Planejamento de Produção

**Automação em VBA que distribui um plano de produção no calendário — respeitando capacidade, tempo de cura e prioridade.**

![VBA](https://img.shields.io/badge/VBA-2C5FA6?style=for-the-badge&logo=microsoft&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)

</div>

---

> 💡 Projeto de demonstração técnica, com dados e parâmetros fictícios.

> 🏭 Este projeto simula um sistema de planejamento de produção (APS), inspirado em desafios reais de nivelamento de capacidade e priorização que venho enfrentando na minha atuação profissional em ambientes industriais.

## 📌 O problema

A programação manual, item por item, gerava dores recorrentes:

| ❌ Antes | ✅ Depois (Mini-APS) |
|---|---|
| Carga empilhada nos primeiros dias do calendário | Curva de produção plana, nivelada automaticamente |
| Estouro de capacidade sem aviso | Múltiplos tetos checados em tempo real |

## 🚀 Diferenciais técnicos

| | Recurso | O que faz |
|---|---|---|
| 🌊 | **Valley-Fill anti-cascata** | Sempre aloca no dia com menor carga do momento (não em ordem cronológica), recalculando a cada lote. Roda em 2 fases: nivelamento até um alvo diário, depois pente-fino nas sobras até o teto. |
| 🔮 | **Forward Window Checking** | Como a alocação pula dias, o sistema verifica se o molde fica livre durante *toda* a janela de cura à frente — não só no dia da alocação — evitando colisão com a cura de outro lote. |
| 🎯 | **Score de prioridade** | Ordena por urgência → tipo de recurso (máquina) → fator de gargalo (`plano ÷ moldes disponíveis × leadtime`) → volume. Itens mais restritos entram primeiro, enquanto ainda há espaço. |
| 🟩 | **Células fixas (fundo verde)** | Ordens travadas manualmente na planilha são reconhecidas, preservadas e descontadas do saldo — o plano se adapta *ao redor* delas, sem sobrepor. |
| 🧱 | **Múltiplos tetos simultâneos** | Fábrica, máquina, 4 categorias de produto e capacidade individual por item — o menor espaço disponível sempre prevalece. |
| 📦 | **Lote mínimo estrito** | Evita fracionamento excessivo, com exceção só quando o próprio plano/moldes já é menor que o mínimo. |
| 🧾 | **Log de motivo residual** | Quando um item não é 100% alocado, o motivo exato é gravado automaticamente (teto atingido, falta de molde, sem dias disponíveis). |
| 📊 | **Resumo executivo** | Ao final, mostra total alocado, saldo residual, média diária, dia mais alto/baixo e dias acima do teto. |

## 🔄 Fluxo de execução

```mermaid
flowchart LR
    A[🟩 Fase 0\nTrava células fixas] --> B[🌊 Fase 1\nNivela até o alvo diário]
    B --> C[📦 Fase 2\nPente-fino nas sobras]
    C --> D[🧾 Grava resultado\n+ log + resumo]
```

## 🛠️ Tecnologias

`VBA` · `Excel` · estrutura de dados tipada · forward window checking

---

<div align="center">

📁 Projeto de portfólio com dados fictícios · Todos os direitos reservados ao autor

</div>
