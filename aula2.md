# Slide 1: DIFERENÇA ENTRE IA REATIVA E AGENTES INTELIGENTES
## Do mapeamento passivo à autonomia e execução orientada a objetivos.

### Comparativo: Analogia Termostato vs. Climatização
- **Termostato Comum**
  - Resposta direta e inflexível.
  - Apenas desliga ao atingir uma temperatura específica.
  - Atua sobre um único estímulo.
- **Climatização Inteligente**
  - Ajuste proativo.
  - Monitora previsão do tempo, ocupação do recinto e incidência solar para antecipar o conforto térmico.

### O Básico: Inteligência Artificial Reativa
- Ausência de memória de longo prazo nativa.
- Mapeamento direto de percepções imediatas para ações (`Prompt in → Text out`).
- Abordagem zero-shot: gera respostas sem intenção ou avaliação de impacto real.
- **O Problema:** Soluções frágeis. Se o ambiente técnico muda, a IA falha e exige supervisão humana constante.

<!-- SLIDE -->

# Slide 2: O SALTO EVOLUTIVO
## 01. Estado
- A capacidade de manter o histórico contínuo da operação e o contexto ambiental além de uma janela de conversa isolada.

## 02. Objetivo
- O sistema não se limita a fornecer uma resposta textual; ele dedica-se a alcançar proativamente um estado final predeterminado pelo usuário.

### A Anatomia do Agente
- **Sensores (Percepção):** Integrações diretas a fontes de dados em tempo real e APIs de sistemas.
- **Atuadores (Intervenção):** Capacidade técnica para executar códigos lógicos, realizar cálculos, gravar arquivos e enviar e-mails.

### O Motor da Autonomia
- **Sentir:** Coletar dados do ambiente via sensores.
- **Pensar:** Processar o estado atual em relação ao objetivo final.
- **Agir:** Utilizar atuadores para modificar o ambiente.

> *A grande revolução não está no tamanho do modelo gerador, mas sim em sua inserção neste fluxo estruturado contínuo.*

<!-- SLIDE -->

# Slide 3: O FIM DA FRAGILIDADE
## A Resiliência Sistêmica elimina o gargalo da gestão de indisponibilidades.

### Fluxo de Falha: IA Reativa vs. Agente Inteligente
| Etapa | IA Reativa | Agente Inteligente |
|-------|------------|-------------------|
| **Falha** | Encontra erro 500 de servidor | Encontra falha no sistema |
| **Reação** | Desiste e repassa erro ao usuário | Aguarda ou busca rota alternativa |
| **Resultado** | Processo interrompido | Garante a entrega do valor |
| **Rastreabilidade** | Nenhuma | Documenta o próprio raciocínio |

<!-- SLIDE -->

# Slide 4: O LOOP DE OBSERVAÇÃO
## Conceito de Persistência Autônoma
- O agente avalia constantemente se o resultado de sua última ação o aproximou ou o afastou do objetivo final da missão.

### Lógica de Execução (Pseudocódigo)
```python
while estado_atual != 'sucesso' and tentativas < 3:
    resultado = realizar_acao(objetivo)
    if validar_resultado(resultado):
        break
    else:
        tentativas += 1
        print("Ajustando estratégia...")