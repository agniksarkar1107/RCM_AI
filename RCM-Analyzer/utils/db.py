import chromadb
import os
import uuid
import json
from typing import Dict, List, Any, Union
import logging
import sys
# Use built-in sqlite3 only
import sqlite3
import importlib.util
import platform

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def initialize_chroma(collection_name: str):
    """
    Initialize a ChromaDB collection
    
    Args:
        collection_name: Name of the collection to create/load
        
    Returns:
        ChromaDB collection instance
    """
    try:
        # Create a persistent client
        persist_directory = os.path.join(os.getcwd(), "chroma_db")
        os.makedirs(persist_directory, exist_ok=True)
        
        # Check SQLite version and use in-memory if needed
        sqlite_version = sqlite3.sqlite_version_info
        min_version_required = (3, 35, 0)
        
        if sqlite_version < min_version_required:
            logger.warning(f"SQLite version {sqlite_version} is below required version {min_version_required}. Using in-memory database instead.")
            # Use in-memory database as fallback
            client = chromadb.Client()
            logger.info("Using in-memory ChromaDB client due to SQLite version constraints")
        else:
            # Use persistent storage if SQLite version is sufficient
            client = chromadb.PersistentClient(path=persist_directory)
            logger.info(f"Using persistent ChromaDB client at {persist_directory}")
        
        # Get or create collection
        try:
            # Try with sentence-transformers embedding function
            if importlib.util.find_spec("sentence_transformers"):
                from chromadb.utils.embedding_functions import SentenceTransformerEmbeddingFunction
                embedding_function = SentenceTransformerEmbeddingFunction(model_name="all-MiniLM-L6-v2")
                collection = client.get_or_create_collection(name=collection_name, embedding_function=embedding_function)
                logger.info(f"Created/loaded collection {collection_name} with sentence-transformers")
            else:
                # Fall back to creating collection without embedding function
                collection = client.get_or_create_collection(name=collection_name)
                logger.info(f"Created/loaded collection {collection_name} without embedding function")
        except Exception as embedding_error:
            logger.warning(f"Failed to create collection with embedding function: {str(embedding_error)}")
            # Final fallback - no embedding function
            collection = client.get_or_create_collection(name=collection_name)
            logger.info(f"Created/loaded collection {collection_name} without embedding function")
        
        return collection
    except Exception as e:
        logger.error(f"Failed to initialize ChromaDB: {str(e)}")
        raise e

def store_in_chroma(collection, data: Dict[str, Any]):
    """
    Store data in ChromaDB collection
    
    Args:
        collection: ChromaDB collection
        data: Data to store
    """
    try:
        # Process different types of data for storage
        if "raw_text" in data and data["raw_text"]:
            # For PDF or DOCX, store the extracted text in chunks
            if "extracted_text" in data:
                text = data["extracted_text"]
                # Split text into chunks of approximately 500 tokens
                chunks = split_text_into_chunks(text, chunk_size=1000)
                
                # Store each chunk
                ids = []
                documents = []
                metadatas = []
                
                for i, chunk in enumerate(chunks):
                    chunk_id = str(uuid.uuid4())
                    ids.append(chunk_id)
                    documents.append(chunk)
                    metadatas.append({
                        "chunk_index": i,
                        "source": data["metadata"]["file_name"],
                        "file_type": data["metadata"]["file_type"],
                        "total_chunks": len(chunks)
                    })
                
                # Add documents to collection
                try:
                    collection.add(
                        ids=ids,
                        documents=documents,
                        metadatas=metadatas
                    )
                    logger.info(f"Stored {len(chunks)} text chunks in ChromaDB")
                except Exception as add_error:
                    logger.warning(f"Error adding documents to ChromaDB: {str(add_error)}.")
                    logger.info("Proceeding with analysis without storing in ChromaDB")
        
        else:
            # For structured data from Excel or CSV
            ids = []
            documents = []
            metadatas = []
            
            # Store control objectives
            for i, objective in enumerate(data.get("control_objectives", [])):
                obj_id = str(uuid.uuid4())
                ids.append(obj_id)
                
                # Create a text representation of the objective
                doc_text = f"Department: {objective.get('department', '')}\n"
                doc_text += f"Control Objective: {objective.get('objective', '')}\n"
                doc_text += f"What Can Go Wrong: {objective.get('what_can_go_wrong', '')}\n"
                doc_text += f"Risk Level: {objective.get('risk_level', '')}\n"
                doc_text += f"Control Activities: {objective.get('control_activities', '')}\n"
                
                if objective.get("is_gap", False):
                    doc_text += f"Gap Details: {objective.get('gap_details', '')}\n"
                    doc_text += f"Proposed Control: {objective.get('proposed_control', '')}\n"
                
                documents.append(doc_text)
                
                # Create metadata
                metadata = {
                    "document_type": "control_objective",
                    "department": objective.get("department", ""),
                    "risk_level": objective.get("risk_level", ""),
                    "has_gap": objective.get("is_gap", False),
                    "source": data["metadata"]["file_name"]
                }
                metadatas.append(metadata)
            
            # Store gaps
            for i, gap in enumerate(data.get("gaps", [])):
                gap_id = str(uuid.uuid4())
                ids.append(gap_id)
                
                # Create a text representation of the gap
                doc_text = f"Department: {gap.get('department', '')}\n"
                doc_text += f"Control Objective: {gap.get('control_objective', '')}\n"
                doc_text += f"Gap Title: {gap.get('gap_title', '')}\n"
                doc_text += f"Description: {gap.get('description', '')}\n"
                doc_text += f"Risk Impact: {gap.get('risk_impact', '')}\n"
                doc_text += f"Proposed Solution: {gap.get('proposed_solution', '')}\n"
                
                documents.append(doc_text)
                
                # Create metadata
                metadata = {
                    "document_type": "gap",
                    "department": gap.get("department", ""),
                    "source": data["metadata"]["file_name"]
                }
                metadatas.append(metadata)
            
            # Add documents to collection if any
            if ids:
                try:
                    collection.add(
                        ids=ids,
                        documents=documents,
                        metadatas=metadatas
                    )
                    logger.info(f"Stored {len(ids)} documents in ChromaDB")
                except Exception as add_error:
                    logger.warning(f"Error adding documents to ChromaDB: {str(add_error)}.")
                    logger.info("Proceeding with analysis without storing in ChromaDB") 
    
    except Exception as e:
        logger.error(f"Error storing data in ChromaDB: {str(e)}")
        # Don't raise the exception, just log it and continue
        logger.info("Proceeding with analysis without storing in ChromaDB")

def query_chroma(collection, query: str, n_results: int = 5, filter_dict: Dict = None):
    """
    Query the ChromaDB collection
    
    Args:
        collection: ChromaDB collection
        query: Query string
        n_results: Number of results to return
        filter_dict: Filter for query
        
    Returns:
        Query results
    """
    try:
        results = collection.query(
            query_texts=[query],
            n_results=n_results,
            where=filter_dict
        )
        
        return results
    
    except Exception as e:
        logger.error(f"Error querying ChromaDB: {str(e)}")
        raise

def split_text_into_chunks(text: str, chunk_size: int = 1000, overlap: int = 100) -> List[str]:
    """
    Split text into overlapping chunks
    
    Args:
        text: Text to split
        chunk_size: Approximate size of each chunk
        overlap: Number of characters to overlap between chunks
        
    Returns:
        List of text chunks
    """
    if not text:
        return []
    
    chunks = []
    start = 0
    text_length = len(text)
    
    while start < text_length:
        # Find the end of this chunk
        end = min(start + chunk_size, text_length)
        
        # If we're not at the end of the text, try to find a good breaking point
        if end < text_length:
            # Look for a newline or period near the end
            last_newline = text.rfind('\n', start, end)
            last_period = text.rfind('.', start, end)
            
            # Use the later of the two if they're within a reasonable distance of the end
            if last_newline > start + (chunk_size // 2):
                end = last_newline + 1  # Include the newline
            elif last_period > start + (chunk_size // 2):
                end = last_period + 1  # Include the period
        
        # Extract the chunk
        chunk = text[start:end]
        chunks.append(chunk)
        
        # Move to the next chunk with overlap
        start = end - overlap if end < text_length else text_length
    
    return chunks 