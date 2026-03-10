# Gerador de Variantes de SKU e Expansão de Hierarquia (VBA)

## 📌 Sobre o Projeto

Este projeto automatiza o processo de criação de códigos de variantes a partir de um SKU mestre (Pai). A ferramenta cruza informações entre uma interface de entrada e um banco de dados de dimensões para "abrir" a hierarquia de produtos, gerando automaticamente os códigos de variantes, vinculando preços e setores.

### 💡 Por que este projeto é importante?

No varejo e na logística, um único produto pode ter dezenas de variações. Criar esses códigos manualmente é suscetível a erros de digitação e inconsistências de dados. Esta automação garante:
* **Integridade de Dados:** Valida se o código mestre possui exatamente 6 dígitos antes de processar.
* **Agilidade no Cadastro:** Transforma uma lista de SKUs mestre em uma lista completa de variantes (Pai + Sufixo 001, 002, etc.) em segundos.
* **Sincronização Automática:** Utiliza lógica de `VLOOKUP` via VBA para buscar o setor correspondente no dataset de dimensões.
* **Padronização Visual:** Aplica formatação condicional e estética (bordas vermelhas, cabeçalhos destacados) para facilitar a conferência imediata.

---

## 🛡️ Disclaimer e Privacidade (Dados Anonimizados)

* **Privacidade:** Todos os dados de SKUs, preços e setores apresentados nas imagens e no código foram **totalmente anonimizados**. Os valores e descrições são fictícios, criados apenas para demonstrar a funcionalidade da lógica de automação.

---

## 🏗️ Estrutura e Funcionamento

A ferramenta opera conectando duas abas principais:

### 1. Dataset de Dimensões (`Dataset_Dimensoes`)
Funciona como o "Cadastro Mestre", contendo a relação de SKUs e seus respectivos setores. É a fonte de consulta para a validação de quantas variantes cada produto possui.

<img width="482" height="829" alt="dataset" src="https://github.com/user-attachments/assets/4c9e2f3a-408d-4c00-bbb3-1ed25289c1df" />

*Figura 1: Base de dados de dimensões para consulta.*

### 2. Interface Geral (`Geral`)
É onde o usuário insere os SKUs que deseja expandir. O algoritmo limpa as colunas de saída, valida a estrutura dos códigos de entrada e gera a nova tabela formatada.

<img width="1343" height="836" alt="geral" src="https://github.com/user-attachments/assets/44425feb-b019-482c-b4c0-fd8b1c4a11d5" />

*Figura 2: Painel de controle e resultado da expansão.*

---

## 💻 Destaques do Código VBA

O script `GerarCodVar` utiliza técnicas avançadas para garantir performance e precisão:

* **Otimização de Performance:** Uso de `Application.ScreenUpdating = False` para evitar o "piscar" da tela e acelerar o processamento.
* **Lógica de Expansão:** Se um produto possui mais de uma variante, o código inicia um loop para gerar sufixos formatados (ex: `001`, `002`) usando a função `Format(j, "000")`.
* **Tratamento de Erros:** Exibe uma `MsgBox` de alerta caso um código não siga o padrão de 6 dígitos.
* **Automação Estética:** O código define bordas grossas na cor vermelha (`Color:=vbRed`), negrito nos cabeçalhos e preenchimento amarelo, além de ajustar automaticamente a largura das colunas (`AutoFit`).
