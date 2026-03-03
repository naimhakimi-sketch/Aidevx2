-- Enable pgvector extension
create extension if not exists vector
with
  schema extensions;

-- Create document_chunks table
create table document_chunks (
  id uuid primary key default gen_random_uuid(),
  project_id uuid references public.projects (id) on delete cascade not null,
  document_path text not null,
  content text not null,
  embedding extensions.vector(1536) -- OpenAI embeddings have 1536 dimensions
);

-- Protect the table with RLS
alter table public.document_chunks enable row level security;

-- Policies for document_chunks
create policy "Users can view chunks for their projects"
on public.document_chunks for select
using (
  auth.uid() = (select user_id from public.projects where id = public.document_chunks.project_id)
);

create policy "Users can insert chunks for their projects"
on public.document_chunks for insert
with check (
  auth.uid() = (select user_id from public.projects where id = public.document_chunks.project_id)
);

create policy "Users can update chunks for their projects"
on public.document_chunks for update
using (
  auth.uid() = (select user_id from public.projects where id = public.document_chunks.project_id)
);

create policy "Users can delete chunks for their projects"
on public.document_chunks for delete
using (
  auth.uid() = (select user_id from public.projects where id = public.document_chunks.project_id)
);

-- Create a function to similarity search for document chunks
create or replace function match_document_chunks (
  query_embedding extensions.vector(1536),
  match_threshold float,
  match_count int,
  p_project_id uuid
)
returns table (
  id uuid,
  project_id uuid,
  document_path text,
  content text,
  similarity float
)
language sql stable
set search_path = public, extensions
as $$
  select
    document_chunks.id,
    document_chunks.project_id,
    document_chunks.document_path,
    document_chunks.content,
    1 - (document_chunks.embedding <=> query_embedding) as similarity
  from document_chunks
  where document_chunks.project_id = p_project_id
  and 1 - (document_chunks.embedding <=> query_embedding) > match_threshold
  order by document_chunks.embedding <=> query_embedding
  limit match_count;
$$;
