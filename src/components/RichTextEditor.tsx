import { useEditor, EditorContent, Editor } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import Placeholder from '@tiptap/extension-placeholder'
import {
    Bold, Italic, Strikethrough, Underline as UnderlineIcon,
    Heading1, Heading2, Heading3,
    List, ListOrdered, Quote, Undo, Redo
} from 'lucide-react'
import { useEffect, useState } from 'react'

interface RichTextEditorProps {
    content: string;
    onChange: (html: string) => void;
    placeholder?: string;
    editable?: boolean;
}

const MenuBar = ({ editor, visible }: { editor: Editor | null, visible: boolean }) => {
    if (!editor) {
        return null
    }

    const buttons = [
        {
            icon: <Bold size={14} />,
            title: 'Bold',
            action: () => editor.chain().focus().toggleBold().run(),
            isActive: editor.isActive('bold'),
        },
        {
            icon: <Italic size={14} />,
            title: 'Italic',
            action: () => editor.chain().focus().toggleItalic().run(),
            isActive: editor.isActive('italic'),
        },
        {
            icon: <UnderlineIcon size={14} />,
            title: 'Underline',
            action: () => editor.chain().focus().toggleUnderline().run(),
            isActive: editor.isActive('underline'),
        },
        {
            icon: <Strikethrough size={14} />,
            title: 'Strike',
            action: () => editor.chain().focus().toggleStrike().run(),
            isActive: editor.isActive('strike'),
        },
        { type: 'divider' },
        {
            icon: <Heading1 size={14} />,
            title: 'Heading 1',
            action: () => editor.chain().focus().toggleHeading({ level: 1 }).run(),
            isActive: editor.isActive('heading', { level: 1 }),
        },
        {
            icon: <Heading2 size={14} />,
            title: 'Heading 2',
            action: () => editor.chain().focus().toggleHeading({ level: 2 }).run(),
            isActive: editor.isActive('heading', { level: 2 }),
        },
        {
            icon: <Heading3 size={14} />,
            title: 'Heading 3',
            action: () => editor.chain().focus().toggleHeading({ level: 3 }).run(),
            isActive: editor.isActive('heading', { level: 3 }),
        },
        { type: 'divider' },
        {
            icon: <List size={14} />,
            title: 'Bullet List',
            action: () => editor.chain().focus().toggleBulletList().run(),
            isActive: editor.isActive('bulletList'),
        },
        {
            icon: <ListOrdered size={14} />,
            title: 'Ordered List',
            action: () => editor.chain().focus().toggleOrderedList().run(),
            isActive: editor.isActive('orderedList'),
        },
        {
            icon: <Quote size={14} />,
            title: 'Blockquote',
            action: () => editor.chain().focus().toggleBlockquote().run(),
            isActive: editor.isActive('blockquote'),
        },

        { type: 'divider' },
        {
            icon: <Undo size={14} />,
            title: 'Undo',
            action: () => editor.chain().focus().undo().run(),
            disabled: !editor.can().chain().focus().undo().run(),
        },
        {
            icon: <Redo size={14} />,
            title: 'Redo',
            action: () => editor.chain().focus().redo().run(),
            disabled: !editor.can().chain().focus().redo().run(),
        },
    ]

    return (
        <div className={`absolute -top-10 left-0 flex flex-wrap gap-0.5 p-1 bg-white border border-slate-200 shadow-md rounded-lg z-20 transition-all duration-200 origin-bottom-left ${visible ? 'opacity-100 scale-100 translate-y-0' : 'opacity-0 scale-95 translate-y-2 pointer-events-none'}`}>
            {buttons.map((btn, index) => {
                if (btn.type === 'divider') {
                    return <div key={index} className="w-px h-5 bg-slate-200 mx-1 self-center" />
                }
                return (
                    <button
                        key={index}
                        onClick={btn.action}
                        disabled={btn.disabled}
                        className={`p-1.5 rounded text-slate-500 hover:bg-slate-100 hover:text-slate-900 transition-colors ${btn.isActive ? 'bg-slate-100 text-slate-900 font-bold' : ''} ${btn.disabled ? 'opacity-30 cursor-not-allowed' : ''}`}
                        title={btn.title}
                        type="button"
                    >
                        {btn.icon}
                    </button>
                )
            })}
        </div>
    )
}

const RichTextEditor = ({ content, onChange, placeholder = 'Start typing...', editable = true }: RichTextEditorProps) => {
    const [isHovered, setIsHovered] = useState(false);

    const editor = useEditor({
        extensions: [
            StarterKit.configure({
                heading: {
                    levels: [1, 2, 3],
                },
            }),
            Placeholder.configure({
                placeholder,
            }),
        ],
        content: content,
        editable: editable,
        onUpdate: ({ editor }) => {
            onChange(editor.getHTML())
        },
        editorProps: {
            attributes: {
                class: 'prose prose-slate max-w-none focus:outline-none min-h-[1.5em] prose-p:text-[11px] prose-p:leading-relaxed prose-p:m-0 prose-p:mb-1.5 prose-li:text-[11px] prose-ul:my-1 prose-ol:my-1 prose-h1:text-base prose-h1:font-bold prose-h2:text-sm prose-h2:font-bold prose-h3:text-[11px] prose-h3:font-bold prose-hr:my-2',
            },
        },
    })

    // Update content if external content changes significantly (though risky with local state)
    // Generally Tiptap manages its own state, but if we switch sections (new content prop), we need to update
    useEffect(() => {
        if (editor && content !== editor.getHTML()) {
            // Only update if not currently focused to avoid destroying the caret position while typing
            if (!editor.isFocused) {
                editor.commands.setContent(content)
            }
        }
    }, [content, editor])

    const isVisible = editor?.isFocused || isHovered;

    return (
        <div
            className="bg-transparent flex flex-col relative group/editor w-full"
            onMouseEnter={() => setIsHovered(true)}
            onMouseLeave={() => setIsHovered(false)}
        >
            <MenuBar editor={editor} visible={isVisible} />
            <div className={`transition-colors duration-200 border-l-[3px] py-1 ${isVisible ? 'border-slate-200 pl-4 -ml-4' : 'border-transparent pl-4 -ml-4'}`}>
                <EditorContent editor={editor} className="flex-1" />
            </div>
        </div>
    )
}

export default RichTextEditor
