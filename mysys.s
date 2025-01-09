; hello.asm
[BITS 16]          ; Configuração para 16 bits (modo real)
[ORG 0x7C00]       ; Define o endereço de carregamento

start:
    jmp cons
msg db "my system version 1.1 put you message here....", 0 ; Mensagem a ser impressa (terminada em 0)

times 510-($-$$) db 0 ; Preenche até 510 bytes
dw 0xAA55             ; Assinatura do setor de boot (0xAA55)
cons:
    push ds
    mov ax,0xb800
    mov ds,ax
    mov cx,4000
    mov ax,0x6020
    mov bx,0
loop1:
    ds
    mov [bx],ax
    inc bx
    inc bx
    dec cx
    cmp cx,0
    jnz loop1

    pop ds


printss:
    mov ax,0x1000
    mov ax, area1     ; Ponteiro para a mensagem
    add ax,0
    mov bp,ax
prints:
    push bp
    mov bl,0x60
    mov bh,0
    mov dh,0
    mov dl,0
    mov cx,1024
    mov ah,0x13
    mov al,0x1
    int 0x10
    mov ax,0
    int 0x16
    pop ax
    add ax,1024
    mov bp,ax
    jmp prints
.hang:
    cli             ; Desabilita interrupções
    hlt             ; Entra em estado de espera
error:
    mov bp,err
    mov bl,0x60
    mov bh,0
    mov dh,0
    mov dl,0
    mov cx,32
    mov ah,0x13
    mov al,0x1
    int 0x10
    mov ax,0
    int 0x16
.hangs:
    cli             ; Desabilita interrupções
    hlt             ; Entra em estado de espera
hexs:
db "0123456789abcdf"
err:
db "error......................................................................................."
area0:
db 0x10,0x0
dw 32
dw 0x0,0x1000
dw 0
dw 0
dw 0
times 32 db 0 ; Preenche até 510 bytes
area2:
area1:
db "#file1.txt| ram drive text file system\n\n\n#file2.txt|ram drive ver 1.1\n",0
times 1024 db 0 ; Preenche até 510 bytes