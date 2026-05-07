from planner.core.vector_store import build_vector_index, vector_index_status


if __name__ == "__main__":
    index = build_vector_index(force=True)
    status = vector_index_status()
    print("向量知识库重建完成")
    print(status)
