#!/bin/bash

# 📤 Script de Push para GitHub
# Simplifica: git status, add, commit, push

echo "================================================"
echo "PUSH PARA GITHUB - SOMA"
echo "================================================"
echo ""

# 1. Ver status
echo "📊 Status dos ficheiros:"
git status --short
echo ""

# 2. Perguntar se quer adicionar tudo
read -p "Adicionar TODAS as alterações? (s/n) " -n 1 -r
echo ""

if [[ $REPLY =~ ^[Ss]$ ]]; then
    echo "📝 Adicionando alterações..."
    git add .

    # 3. Ver o que foi staged
    echo ""
    echo "✅ Ficheiros staged:"
    git diff --cached --name-only
    echo ""

    # 4. Pedir mensagem de commit
    read -p "Mensagem de commit: " commit_msg

    if [ -z "$commit_msg" ]; then
        echo "❌ Mensagem vazia!"
        exit 1
    fi

    # 5. Fazer commit
    echo ""
    echo "📦 Criando commit..."
    git commit -m "$commit_msg"

    # 6. Fazer push
    echo ""
    echo "🚀 Enviando para GitHub..."
    git push origin main

    echo ""
    echo "================================================"
    echo "✅ PUSH CONCLUÍDO!"
    echo "================================================"
else
    echo "❌ Operação cancelada"
    exit 1
fi
