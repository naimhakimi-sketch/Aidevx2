import { useState, useEffect, useCallback } from 'react';
import { supabase } from '../lib/supabase';
import { useAuth } from '../context/AuthContext';

export interface DocComment {
    id: string;
    docId: string;
    projectId: string;
    sectionIndex: number;
    parentId: string | null;
    authorId: string;
    authorName: string;
    authorEmail: string;
    content: string;
    resolved: boolean;
    resolvedBy: string | null;
    createdAt: string;
    replies?: DocComment[];
}

export function useDocumentComments(docId: string | undefined, projectId: string | undefined) {
    const { user } = useAuth();
    const [comments, setComments] = useState<DocComment[]>([]);
    const [loading, setLoading] = useState(false);

    const fetchComments = useCallback(async () => {
        if (!docId || !projectId) return;

        setLoading(true);
        try {
            const { data, error } = await supabase
                .from('doc_comments')
                .select('*')
                .eq('doc_id', docId)
                .eq('project_id', projectId)
                .order('created_at', { ascending: true });

            if (error) throw error;

            // Fetch author profiles separately since author_id references auth.users, not profiles
            let profilesMap = new Map<string, any>();
            if (data && data.length > 0) {
                const authorIds = [...new Set(data.map((c: any) => c.author_id).filter(Boolean))];
                if (authorIds.length > 0) {
                    const { data: profilesData } = await supabase
                        .from('profiles')
                        .select('id, email, full_name')
                        .in('id', authorIds);
                    
                    profilesMap = new Map(profilesData?.map(p => [p.id, p]) || []);
                }
            }

            const flat: DocComment[] = (data || []).map((c: any) => {
                const profile = profilesMap.get(c.author_id) || { email: '', full_name: '' };
                return {
                    id: c.id,
                    docId: c.doc_id,
                    projectId: c.project_id,
                    sectionIndex: c.section_index,
                    parentId: c.parent_id,
                    authorId: c.author_id,
                    authorName: profile.full_name || '',
                    authorEmail: profile.email || '',
                    content: c.content,
                    resolved: c.resolved,
                    resolvedBy: c.resolved_by,
                    createdAt: c.created_at,
                };
            });

            // Build threaded structure: top-level comments with replies nested
            const topLevel = flat.filter(c => !c.parentId);
            const replies = flat.filter(c => c.parentId);

            topLevel.forEach(comment => {
                comment.replies = replies.filter(r => r.parentId === comment.id);
            });

            setComments(topLevel);
        } catch (error) {
            console.error('Error fetching comments:', error);
            if (error instanceof Error) {
                console.error('Error details:', error.message);
            }
        } finally {
            setLoading(false);
        }
    }, [docId, projectId]);

    useEffect(() => {
        fetchComments();
    }, [fetchComments]);

    // Subscribe to realtime comment inserts
    useEffect(() => {
        if (!docId || !projectId) return;

        const channel = supabase
            .channel(`comments:${docId}`)
            .on(
                'postgres_changes',
                {
                    event: '*',
                    schema: 'public',
                    table: 'doc_comments',
                    filter: `project_id=eq.${projectId}`,
                },
                () => {
                    // Refetch on any change
                    fetchComments();
                }
            )
            .subscribe();

        return () => {
            supabase.removeChannel(channel);
        };
    }, [docId, projectId, fetchComments]);

    const addComment = async (sectionIndex: number, content: string, parentId?: string) => {
        if (!docId || !projectId || !user) return;

        try {
            const { error } = await supabase
                .from('doc_comments')
                .insert({
                    doc_id: docId,
                    project_id: projectId,
                    section_index: sectionIndex,
                    parent_id: parentId || null,
                    author_id: user.id,
                    content,
                });

            if (error) throw error;
            // Realtime will trigger refetch, but also fetch immediately for responsiveness
            await fetchComments();
        } catch (error) {
            console.error('Error adding comment:', error);
        }
    };

    const resolveComment = async (commentId: string) => {
        if (!user) return;

        try {
            const { error } = await supabase
                .from('doc_comments')
                .update({
                    resolved: true,
                    resolved_by: user.id,
                    updated_at: new Date().toISOString(),
                })
                .eq('id', commentId);

            if (error) throw error;
            await fetchComments();
        } catch (error) {
            console.error('Error resolving comment:', error);
        }
    };

    const deleteComment = async (commentId: string) => {
        try {
            const { error } = await supabase
                .from('doc_comments')
                .delete()
                .eq('id', commentId);

            if (error) throw error;
            await fetchComments();
        } catch (error) {
            console.error('Error deleting comment:', error);
        }
    };

    // Get comment counts per section
    const getCommentCountBySection = useCallback(() => {
        const counts: Record<number, number> = {};
        comments.forEach(c => {
            counts[c.sectionIndex] = (counts[c.sectionIndex] || 0) + 1;
        });
        return counts;
    }, [comments]);

    return {
        comments,
        loading,
        addComment,
        resolveComment,
        deleteComment,
        getCommentCountBySection,
        refreshComments: fetchComments,
    };
}
